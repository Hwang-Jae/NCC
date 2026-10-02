import os
import streamlit as st
import clickhouse_connect
import ollama
import pdfplumber
from docx import Document
import openpyxl
from webdav3.client import Client as WebDavClient
import ctypes
import string
import email.utils
import tempfile

# --- 1. 설정 정보 ---
EMBED_MODEL = 'nomic-embed-text'  # 벡터 변환용 임베딩 모델 (Ollama에서 미리 pull 받아주세요)
LLM_MODEL = 'gpt-oss:20b'        # 답변 생성용 LLM
CLOUDDOC_HOST = 'http://idisk.nvhkorea.com:80'

# --- 2. ClickHouse 연결 설정 ---
def get_db_client():
    return clickhouse_connect.get_client(
        host='localhost', port=8123, username='default', password='clickhouse',
        connect_timeout=300, send_receive_timeout=3000
    )

# DB 테이블 초기화 (file_mtime 필드 추가로 증분 동기화 구현)
def init_clickhouse():
    client = get_db_client()
    client.command("""
        CREATE TABLE IF NOT EXISTS doc_chunks (
            id String,
            drive_name String,
            file_name String,
            file_path String,
            content String,
            embedding Array(Float32),
            file_mtime Float64
        ) ENGINE = MergeTree() ORDER BY id
    """)

# --- 3. 동적 가상 드라이브 감지 기능 ---
def detect_local_drives():
    """Windows PC에 마운트된 클라우드독 가상 드라이브 목록을 자동 스캔합니다."""
    detected = {}
    try:
        # Windows 시스템 드라이브 비트마스크 가져오기
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        for letter in string.ascii_uppercase:
            if bitmask & 1:
                drive_path = f"{letter}:\\"
                # C 드라이브 및 플로피/CD 등 기본 장치 제외
                if letter not in ['A', 'B', 'C']:
                    try:
                        if os.path.exists(drive_path):
                            volumeNameBuffer = ctypes.create_unicode_buffer(1024)
                            ctypes.windll.kernel32.GetVolumeInformationW(
                                ctypes.c_wchar_p(drive_path),
                                volumeNameBuffer,
                                1024,
                                None, None, None, None, 0
                            )
                            volume_name = volumeNameBuffer.value
                            # 볼륨 이름이 있으면 매핑, 없으면 드라이브 문자로 표시
                            label = f"{letter}_{volume_name}" if volume_name else f"{letter}_개인또는공유함"
                            detected[label] = drive_path
                    except Exception:
                        if os.path.exists(drive_path):
                            detected[f"{letter}_공유디스크"] = drive_path
            bitmask >>= 1
    except Exception:
        # 윈도우 API 실패 또는 타 운영체제 환경일 시 기본값 수동 체크 예외 처리
        for letter in ["R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            path = f"{letter}:\\"
            if os.path.exists(path):
                detected[f"{letter}_공유디스크"] = path
    return detected

def fetch_webdav_drives(username, password):
    """사용자가 웹 접속 시 할당받은 클라우드독 최상위 공유 문서함 폴더 목록을 가져옵니다."""
    options = {
        'webdav_url': CLOUDDOC_HOST,
        'webdav_username': username,
        'webdav_password': password
    }
    detected = {}
    try:
        webdav_client = WebDavClient(options)
        items = webdav_client.list()
        for item in items:
            # 폴더 형태의 항목만 추출해 드라이브처럼 인식시킴
            if item.endswith('/') and item != '/':
                folder_name = item.strip('/')
                detected[folder_name] = f"/{folder_name}"
    except Exception as e:
        print(f"WebDAV 드라이브 목록 가져오기 실패: {e}")
    return detected

# --- 4. 클라우드독 WebDAV 로그인 검증 기능 ---
def check_clouddoc_login(username, password):
    options = {
        'webdav_url': CLOUDDOC_HOST,
        'webdav_username': username,
        'webdav_password': password
    }
    try:
        webdav_client = WebDavClient(options)
        webdav_client.list()  # 루트 디렉토리 조회가 정상 작동하면 로그인 성공
        return True
    except Exception as e:
        print(f"로그인 검증 실패: {e}")
        return False

# --- 5. 파일 텍스트 추출 함수들 ---
def extract_text(file_path):
    ext = os.path.splitext(file_path)[-1].lower()
    text = ""
    try:
        if ext == '.pdf':
            with pdfplumber.open(file_path) as pdf:
                text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        elif ext == '.docx':
            doc = Document(file_path)
            text = "\n".join([p.text for p in doc.paragraphs])
        elif ext in ['.xlsx', '.xls']:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            for sheet in wb.sheetnames:
                for row in wb[sheet].iter_rows(values_only=True):
                    text += " ".join([str(cell) for cell in row if cell is not None]) + "\n"
        elif ext == '.txt':
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
    except Exception as e:
        print(f"파일 읽기 실패 ({file_path}): {e}")
    return text

# 텍스트 일정한 크기로 자르기 (Chunking)
def chunk_text(text, chunk_size=600, overlap=100):
    chunks = []
    start = 0
    while start < len(text):
        end = start + chunk_size
        chunks.append(text[start:end])
        start += (chunk_size - overlap)
    return chunks

# WebDAV 환경을 위한 재귀적 파일 조회 헬퍼 함수
def get_webdav_files_recursive(client, path):
    files_list = []
    try:
        items = client.list(path, get_info=True)
        for item in items:
            item_path = item.get('path', '')
            if item_path == path or item_path == path + '/':
                continue
            
            # WebDAV 경로 포맷 통일 및 하위 순회
            if item.get('isdir'):
                files_list.extend(get_webdav_files_recursive(client, item_path))
            else:
                mtime_str = item.get('modified', '')
                try:
                    parsed_date = email.utils.parsedate_to_datetime(mtime_str)
                    mtime_val = parsed_date.timestamp()
                except Exception:
                    mtime_val = 0.0
                
                files_list.append({
                    'path': item_path,
                    'name': os.path.basename(item_path),
                    'mtime': mtime_val
                })
    except Exception as e:
        print(f"WebDAV 디렉토리 순회 중 에러 ({path}): {e}")
    return files_list

# --- 6. 스마트 증분 동기화 (Delta Sync) 통합 구현 ---
def sync_drives_incremental():
    st.info("🔄 최신 변경되거나 누락된 파일만 비교 분석하는 중입니다...")
    client = get_db_client()
    
    # 1. 기존 DB 정보 로드
    db_files = {}
    try:
        result = client.query("SELECT DISTINCT file_path, file_mtime FROM doc_chunks")
        for row in result.result_rows:
            db_files[row[0]] = row[1]
    except Exception as e:
        st.warning(f"기존 색인 정보를 불러올 수 없어 새로 동기화합니다: {e}")

    supported_extensions = ['.pdf', '.docx', '.xlsx', '.xls', '.txt']
    scanned_file_paths = set()
    files_added_or_updated = 0
    files_skipped = 0

    target_drives = st.session_state.target_drives
    drive_mode = st.session_state.drive_mode

    # A. [로컬 가상 드라이브 모드] 동작
    if drive_mode == "local":
        for drive_label, drive_path in target_drives.items():
            st.write(f"🔍 '{drive_label}' 드라이브 직접 동기화 진행 중...")
            for root, dirs, files_in_dir in os.walk(drive_path):
                for file in files_in_dir:
                    ext = os.path.splitext(file)[-1].lower()
                    if ext in supported_extensions:
                        full_path = os.path.join(root, file)
                        scanned_file_paths.add(full_path)
                        
                        try:
                            current_mtime = os.path.getmtime(full_path)
                        except Exception:
                            continue
                        
                        # 파일 내용 수정 판단
                        if full_path in db_files and db_files[full_path] == current_mtime:
                            files_skipped += 1
                            continue
                        
                        if full_path in db_files:
                            client.command(f"ALTER TABLE doc_chunks DELETE WHERE file_path = '{full_path.replace('\\', '\\\\')}'")
                        
                        full_text = extract_text(full_path)
                        if not full_text.strip():
                            continue
                            
                        chunks = chunk_text(full_text)
                        data_to_insert = []
                        for i, chunk in enumerate(chunks):
                            try:
                                response = ollama.embed(model=EMBED_MODEL, input=chunk)
                                embedding = response['embeddings'][0]
                                chunk_id = f"{drive_label}_{file}_{i}_{current_mtime}"
                                data_to_insert.append([
                                    chunk_id, drive_label, file, full_path, chunk, embedding, current_mtime
                                ])
                            except Exception as ex:
                                print(f"임베딩 변환 에러 ({file}): {ex}")
                                continue
                                
                        if data_to_insert:
                            client.insert(
                                'doc_chunks', 
                                data_to_insert, 
                                column_names=['id', 'drive_name', 'file_name', 'file_path', 'content', 'embedding', 'file_mtime']
                            )
                            files_added_or_updated += 1

    # B. [WebDAV 클라우드 연동 모드] 동작 (서버 환경 등으로 가상 드라이브가 없을 때)
    else:
        options = {
            'webdav_url': CLOUDDOC_HOST,
            'webdav_username': st.session_state.username,
            'webdav_password': st.session_state.password
        }
        try:
            webdav_client = WebDavClient(options)
        except Exception as e:
            st.error(f"클라우드독 WebDAV 연결에 실패했습니다: {e}")
            return

        # 임시 다운로드 폴더 생성
        with tempfile.TemporaryDirectory() as temp_dir:
            for drive_label, drive_path in target_drives.items():
                st.write(f"☁️ 클라우드독 '{drive_label}' 폴더 정보 탐색 중...")
                webdav_files = get_webdav_files_recursive(webdav_client, drive_path)
                
                for file_info in webdav_files:
                    remote_path = file_info['path']
                    file_name = file_info['name']
                    current_mtime = file_info['mtime']
                    ext = os.path.splitext(file_name)[-1].lower()
                    
                    if ext in supported_extensions:
                        scanned_file_paths.add(remote_path)
                        
                        # 변경 체크
                        if remote_path in db_files and db_files[remote_path] == current_mtime:
                            files_skipped += 1
                            continue
                            
                        if remote_path in db_files:
                            client.command(f"ALTER TABLE doc_chunks DELETE WHERE file_path = '{remote_path}'")
                        
                        # 파일 일시 다운로드 후 텍스트 파싱
                        local_temp_path = os.path.join(temp_dir, file_name)
                        try:
                            webdav_client.download_file(remote_path, local_temp_path)
                            full_text = extract_text(local_temp_path)
                            os.remove(local_temp_path) # 다운로드 파일 소거
                        except Exception as de:
                            print(f"다운로드 실패 ({remote_path}): {de}")
                            continue
                            
                        if not full_text.strip():
                            continue
                            
                        chunks = chunk_text(full_text)
                        data_to_insert = []
                        for i, chunk in enumerate(chunks):
                            try:
                                response = ollama.embed(model=EMBED_MODEL, input=chunk)
                                embedding = response['embeddings'][0]
                                chunk_id = f"webdav_{drive_label}_{file_name}_{i}_{current_mtime}"
                                data_to_insert.append([
                                    chunk_id, drive_label, file_name, remote_path, chunk, embedding, current_mtime
                                ])
                            except Exception as ex:
                                print(f"임베딩 에러 ({file_name}): {ex}")
                                continue
                                
                        if data_to_insert:
                            client.insert(
                                'doc_chunks', 
                                data_to_insert, 
                                column_names=['id', 'drive_name', 'file_name', 'file_path', 'content', 'embedding', 'file_mtime']
                            )
                            files_added_or_updated += 1

    # 3. 삭제된 파일 동기화 해제
    deleted_files = set(db_files.keys()) - scanned_file_paths
    if deleted_files:
        st.write(f"🧹 공유 디스크에서 지워진 {len(deleted_files)}개의 항목 데이터를 안전하게 정돈하고 있습니다...")
        for del_path in deleted_files:
            escaped_path = del_path.replace('\\', '\\\\')
            client.command(f"ALTER TABLE doc_chunks DELETE WHERE file_path = '{escaped_path}'")

    st.success(
        f"🎉 동기화 마침! [신규/수정 파일]: {files_added_or_updated}건 처리완료, "
        f"[유지 파일]: {files_skipped}건 보존, [삭제 파일]: {len(deleted_files)}건 제거완료."
    )

# --- 7. RAG 검색 및 Ollama 답변 생성 ---
def ask_question(query):
    try:
        query_embed = ollama.embed(model=EMBED_MODEL, input=query)['embeddings'][0]
    except Exception as e:
        return f"임베딩 서버(Ollama) 접속 오류. Ollama가 구동 중인지 및 '{EMBED_MODEL}' 모델이 설치되어 있는지 체크하세요. 에러: {e}", []
    
    client = get_db_client()
    search_query = """
        SELECT content, file_name, file_path, cosineDistance(embedding, {q}) as distance 
        FROM doc_chunks 
        ORDER BY distance ASC 
        LIMIT 4
    """
    
    try:
        result = client.query(search_query, parameters={'q': query_embed})
    except Exception as e:
        return f"ClickHouse 검색 실행 중 에러: {e}", []
    
    context = ""
    sources = set()
    for row in result.result_rows:
        context += f"[참고문서: {row[1]} (경로: {row[2]})]\n{row[0]}\n\n"
        sources.add(f"{row[1]} ({row[2]})")
        
    if not context:
        return "보유하고 있는 클라우드독 공유 드라이브에서 관련 참고 내용을 검색하지 못했습니다.", []

    system_prompt = f"""당신은 사내 문서 정보를 바탕으로 완벽히 답변하는 스마트 인공지능 어시스턴트입니다. 
제시된 [참고문서] 내용만을 철저하게 근거로 삼아 정직하게 답변해 주세요. 
절대로 주관적이거나 문서에 서술되지 않은 내용을 지어내서 답하면 안 됩니다.
자연스럽게 참고 파일명(출처)을 답변 안에서 언급해 주어 신뢰성을 강화해 주세요.

[참고문서]
{context}
"""
    
    try:
        response = ollama.generate(model=LLM_MODEL, prompt=f"{system_prompt}\n\n질문: {query}")
        return response['response'], list(sources)
    except Exception as e:
        return f"Ollama gpt-oss 모델 답변 생성 실패: {e}", list(sources)


# --- 8. Streamlit 멀티유저 대응 웹 UI 화면 구성 ---
st.set_page_config(page_title="공용 클라우드독 AI 지식인", layout="wide")
init_clickhouse()

# 사용자 세션 상태 변수 초기화
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""
if "password" not in st.session_state:
    st.session_state.password = ""
if "messages" not in st.session_state:
    st.session_state.messages = []
if "target_drives" not in st.session_state:
    st.session_state.target_drives = {}
if "drive_mode" not in st.session_state:
    st.session_state.drive_mode = "local" # 'local' or 'webdav'

# --- [화면 1] 로그인 영역 ---
if not st.session_state.logged_in:
    st.markdown("<h2 style='text-align: center; color: #4B9CD3;'>📁 클라우드독 사내 통합 검색 AI 로그인</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: gray;'>클라우드독(문서중앙화) 접속 계정 정보를 이용해 로그인하세요.</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            user_id = st.text_input("아이디 (ID)", placeholder="예: jwhwang1")
            user_pw = st.text_input("비밀번호 (PW)", type="password", placeholder="••••••••")
            login_btn = st.form_submit_button("로그인 및 가상 드라이브 연동", use_container_width=True)
            
            if login_btn:
                if not user_id or not user_pw:
                    st.error("ID와 비밀번호를 모두 입력해 주세요.")
                else:
                    with st.spinner("클라우드독 보안인증 시스템 확인 중..."):
                        if check_clouddoc_login(user_id, user_pw):
                            st.session_state.logged_in = True
                            st.session_state.username = user_id
                            st.session_state.password = user_pw
                            
                            # 1단계: Windows 환경에 마운트된 로컬 가상 드라이브 스캔 시도
                            local_drives = detect_local_drives()
                            if local_drives:
                                st.session_state.target_drives = local_drives
                                st.session_state.drive_mode = "local"
                            else:
                                # 2단계: 로컬 드라이브가 없을 경우 웹(WebDAV) 계정에 직접 할당된 최상위 문서함 스캔 시도
                                webdav_drives = fetch_webdav_drives(user_id, user_pw)
                                st.session_state.target_drives = webdav_drives
                                st.session_state.drive_mode = "webdav"
                                
                            st.success(f"🔓 {user_id}님, 본인 인증이 완료되었습니다!")
                            st.rerun()
                        else:
                            st.error("로그인 실패: 클라우드독 계정 정보가 정확하지 않거나 사내망 연결 상태를 확인하세요.")

# --- [화면 2] 로그인 완료 후 메인 Q&A 시스템 영역 ---
else:
    st.title("📂 사내 클라우드독 지식 Q&A 포털")
    mode_text = "로컬 가상 드라이브 연동 모드" if st.session_state.drive_mode == "local" else "실시간 클라우드 WebDAV 연동 모드"
    st.subheader(f"반갑습니다, {st.session_state.username}님! ({mode_text})")

    # 사이드바 설정
    with st.sidebar:
        st.header("👤 사용자 정보")
        st.info(f"계정: {st.session_state.username}\n상태: 연동 성공\n모드: {st.session_state.drive_mode.upper()}")
        
        if st.button("로그아웃", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.username = ""
            st.session_state.password = ""
            st.session_state.messages = []
            st.session_state.target_drives = {}
            st.rerun()

        st.markdown("---")
        st.header("⚙️ 동적 드라이브 감지 정보")
        st.markdown("**인식된 할당 스페이스:**")
        
        if st.session_state.target_drives:
            for key, path in st.session_state.target_drives.items():
                st.text(f"📁 {key}")
                st.caption(f"경로: {path}")
        else:
            st.warning("접근 권한이 있는 가상 드라이브나 할당 폴더가 비어 있습니다.")
            
        st.markdown("---")
        # 증분 동기화 실행 버튼
        if st.button("🔄 최신 변경정보만 스마트 업데이트", use_container_width=True):
            if st.session_state.target_drives:
                with st.spinner("변경된 대상 분석 중..."):
                    sync_drives_incremental()
            else:
                st.error("동기화할 드라이브 정보가 존재하지 않습니다.")
                
        st.caption("권한이 활성화된 대상 중 새로 생성되거나 변경된 파일 내용만 선별하여 데이터베이스에 신속히 반영합니다.")

    # 메인 채팅 기록 표시
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # 질의 입력 및 RAG 추론 실행
    if user_query := st.chat_input("공유 디스크에 대해 질문하세요 (예: VW 인프라 구성도를 보고 싶어요 / 경주공장 대역 IP 대역은?)"):
        st.session_state.messages.append({"role": "user", "content": user_query})
        with st.chat_message("user"):
            st.markdown(user_query)
            
        with st.chat_message("assistant"):
            with st.spinner("ClickHouse 지식 데이터 분석 중..."):
                answer, source_files = ask_question(user_query)
                
                if source_files:
                    answer += f"\n\n---\n📌 **연관 출처 파일 리스트:**\n" + "\n".join([f"- {src}" for src in source_files])
                    
                st.markdown(answer)
                st.session_state.messages.append({"role": "assistant", "content": answer})