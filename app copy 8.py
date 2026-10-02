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
import time

# --- 1. 설정 정보 ---
EMBED_MODEL = 'nomic-embed-text'  # 벡터 변환용 임베딩 모델 (Ollama에서 미리 pull 받아주세요)
LLM_MODEL = 'gpt-oss:20b'        # 답변 생성용 LLM

CLOUDDOC_HOST = 'https://idisk.nvhkorea.com'
CLOUDDOC_PORT = 443

# --- 2. ClickHouse 연결 설정 ---
def get_db_client():
    return clickhouse_connect.get_client(
        host='localhost', port=8123, username='default', password='clickhouse',
        connect_timeout=30, send_receive_timeout=300
    )

# DB 테이블 초기화 (사용자 식별을 위한 user_id 컬럼 추가 적용)
def init_clickhouse():
    client = get_db_client()
    
    # 1. 신규 환경을 위한 테이블 생성
    client.command("""
        CREATE TABLE IF NOT EXISTS doc_chunks (
            id String,
            user_id String,
            drive_name String,
            file_name String,
            file_path String,
            content String,
            embedding Array(Float32),
            file_mtime Float64
        ) ENGINE = MergeTree() ORDER BY id
    """)
    
    # 2. 기존 테이블 마이그레이션 (user_id가 없을 경우 자동 추가)
    try:
        client.command("ALTER TABLE doc_chunks ADD COLUMN IF NOT EXISTS user_id String DEFAULT 'system'")
        client.command("ALTER TABLE doc_chunks ADD COLUMN IF NOT EXISTS drive_name String DEFAULT ''")
        client.command("ALTER TABLE doc_chunks ADD COLUMN IF NOT EXISTS file_path String DEFAULT ''")
        client.command("ALTER TABLE doc_chunks ADD COLUMN IF NOT EXISTS file_mtime Float64 DEFAULT 0.0")
    except Exception as e:
        print(f"테이블 마이그레이션 예외 발생 (무시 가능): {e}")

# --- 3. 동적 가상 드라이브 감지 기능 ---
def detect_local_drives():
    detected = {}
    try:
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        for letter in string.ascii_uppercase:
            if bitmask & 1:
                drive_path = f"{letter}:\\"
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
                            label = f"{letter}_{volume_name}" if volume_name else f"{letter}_공유함"
                            detected[label] = drive_path
                    except Exception:
                        if os.path.exists(drive_path):
                            detected[f"{letter}_공유디스크"] = drive_path
            bitmask >>= 1
    except Exception:
        for letter in ["R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            path = f"{letter}:\\"
            if os.path.exists(path):
                detected[f"{letter}_공유디스크"] = path
    return detected

def fetch_webdav_drives(username, password):
    options = {
        'webdav_hostname': CLOUDDOC_HOST,
        'webdav_port': CLOUDDOC_PORT,
        'webdav_username': username,
        'webdav_password': password,
        'webdav_root': '/'
    }
    detected = {}
    try:
        webdav_client = WebDavClient(options)
        webdav_client.verify = False
        items = webdav_client.list()
        for item in items:
            if item.endswith('/') and item != '/':
                folder_name = item.strip('/')
                detected[folder_name] = f"/{folder_name}"
    except Exception as e:
        print(f"WebDAV 드라이브 목록 가져오기 실패: {e}")
    return detected

# --- 4. 로그인 검증 ---
def check_clouddoc_login(username, password):
    options = {
        'webdav_hostname': CLOUDDOC_HOST,
        'webdav_port': CLOUDDOC_PORT,
        'webdav_username': username,
        'webdav_password': password,
        'webdav_root': '/'
    }
    try:
        webdav_client = WebDavClient(options)
        webdav_client.verify = False
        webdav_client.list()
        return True
    except Exception as e:
        print(f"최상위 WebDAV 접근 실패: {e}")
        for trial_path in [f"/{username}", "/idisk", "/dav"]:
            try:
                options['webdav_root'] = trial_path
                trial_client = WebDavClient(options)
                trial_client.verify = False
                trial_client.list()
                return True
            except Exception:
                continue
        local_detected = detect_local_drives()
        if local_detected:
            return True
        return False

# --- 5. 텍스트 추출 및 청킹 ---
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

def chunk_text(text, chunk_size=600, overlap=100):
    chunks = []
    start = 0
    while start < len(text):
        end = start + chunk_size
        chunks.append(text[start:end])
        start += (chunk_size - overlap)
    return chunks

def get_webdav_files_recursive(client, path):
    files_list = []
    try:
        items = client.list(path, get_info=True)
        for item in items:
            item_path = item.get('path', '')
            if item_path == path or item_path == path + '/':
                continue
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

# --- 6. 사용자 격리 스마트 증분 동기화 ---
def sync_drives_incremental():
    current_user = st.session_state.username
    st.info(f"🔄 [{current_user}] 계정의 전용 변경 문서만 선별하여 데이터베이스를 동기화 중입니다...")
    client = get_db_client()
    
    # 1. 기존 DB 정보 중 현재 '로그인한 사용자(user_id)'의 데이터만 격리 로드
    db_files = {}
    try:
        result = client.query(
            "SELECT DISTINCT file_path, file_mtime FROM doc_chunks WHERE user_id = %(u)s", 
            parameters={'u': current_user}
        )
        for row in result.result_rows:
            db_files[row[0]] = row[1]
    except Exception as e:
        st.warning(f"기존 색인 정보 로딩 실패: {e}. 초기 연동을 진행합니다.")

    supported_extensions = ['.pdf', '.docx', '.xlsx', '.xls', '.txt']
    scanned_file_paths = set()
    files_added_or_updated = 0
    files_skipped = 0

    target_drives = st.session_state.target_drives
    drive_mode = st.session_state.drive_mode

    # A. [로컬 가상 드라이브 모드]
    if drive_mode == "local":
        for drive_label, drive_path in target_drives.items():
            st.write(f"🔍 '{drive_label}' 스캔 및 사용자 소유권 할당 중...")
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
                        
                        if full_path in db_files and db_files[full_path] == current_mtime:
                            files_skipped += 1
                            continue
                        
                        if full_path in db_files:
                            client.command(
                                f"ALTER TABLE doc_chunks DELETE WHERE file_path = '{full_path.replace('\\', '\\\\')}' AND user_id = '{current_user}'"
                            )
                        
                        full_text = extract_text(full_path)
                        if not full_text.strip():
                            continue
                            
                        chunks = chunk_text(full_text)
                        data_to_insert = []
                        for i, chunk in enumerate(chunks):
                            try:
                                response = ollama.embed(model=EMBED_MODEL, input=chunk)
                                embedding = response['embeddings'][0]
                                chunk_id = f"{current_user}_{drive_label}_{file}_{i}_{current_mtime}"
                                data_to_insert.append([
                                    chunk_id, current_user, drive_label, file, full_path, chunk, embedding, current_mtime
                                ])
                            except Exception as ex:
                                print(f"임베딩 변환 에러 ({file}): {ex}")
                                continue
                                
                        if data_to_insert:
                            client.insert(
                                'doc_chunks', 
                                data_to_insert, 
                                column_names=['id', 'user_id', 'drive_name', 'file_name', 'file_path', 'content', 'embedding', 'file_mtime']
                            )
                            files_added_or_updated += 1

    # B. [WebDAV 클라우드 연동 모드]
    else:
        options = {
            'webdav_hostname': CLOUDDOC_HOST,
            'webdav_port': CLOUDDOC_PORT,
            'webdav_username': st.session_state.username,
            'webdav_password': st.session_state.password,
            'webdav_root': '/'
        }
        try:
            webdav_client = WebDavClient(options)
            webdav_client.verify = False
        except Exception as e:
            st.error(f"클라우드독 WebDAV 연결 실패: {e}")
            return

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
                        
                        if remote_path in db_files and db_files[remote_path] == current_mtime:
                            files_skipped += 1
                            continue
                            
                        if remote_path in db_files:
                            client.command(
                                f"ALTER TABLE doc_chunks DELETE WHERE file_path = '{remote_path}' AND user_id = '{current_user}'"
                            )
                        
                        local_temp_path = os.path.join(temp_dir, file_name)
                        try:
                            webdav_client.download_file(remote_path, local_temp_path)
                            full_text = extract_text(local_temp_path)
                            os.remove(local_temp_path)
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
                                chunk_id = f"webdav_{current_user}_{drive_label}_{file_name}_{i}_{current_mtime}"
                                data_to_insert.append([
                                    chunk_id, current_user, drive_label, file_name, remote_path, chunk, embedding, current_mtime
                                ])
                            except Exception as ex:
                                print(f"임베딩 에러 ({file_name}): {ex}")
                                continue
                                
                        if data_to_insert:
                            client.insert(
                                'doc_chunks', 
                                data_to_insert, 
                                column_names=['id', 'user_id', 'drive_name', 'file_name', 'file_path', 'content', 'embedding', 'file_mtime']
                            )
                            files_added_or_updated += 1

    # 3. 사용자 권한 내에서 삭제된 파일 청크 제거
    deleted_files = set(db_files.keys()) - scanned_file_paths
    if deleted_files:
        st.write(f"🧹 공유 디스크에서 지워진 {len(deleted_files)}개의 옛 데이터를 DB에서 지우는 중...")
        for del_path in deleted_files:
            escaped_path = del_path.replace('\\', '\\\\')
            client.command(
                f"ALTER TABLE doc_chunks DELETE WHERE file_path = '{escaped_path}' AND user_id = '{current_user}'"
            )

    st.success(
        f"🎉 동기화 마침! [신규/수정]: {files_added_or_updated}건, "
        f"[기존 유지]: {files_skipped}건, [삭제 제거]: {len(deleted_files)}건."
    )

# --- 7. 사고 회로 추적형 (RAG Trace) 답변 생성 함수 ---
def ask_question_with_trace(query):
    trace_steps = {}
    start_time = time.time()
    current_user = st.session_state.username
    
    # Step 1: 사용자 질문 분석 및 의도 파악
    trace_steps['step_1_intent'] = {
        "status": "진행 완료",
        "raw_query": query,
        "parsed_keywords": [word for word in query.split() if len(word) > 1]
    }
    
    # Step 2: 임베딩 생성 과정
    try:
        embed_start = time.time()
        query_embed = ollama.embed(model=EMBED_MODEL, input=query)['embeddings'][0]
        embed_elapsed = time.time() - embed_start
        trace_steps['step_2_embed'] = {
            "status": f"성공 ({embed_elapsed:.2f}초 소요)",
            "model_used": EMBED_MODEL,
            "vector_dimension": len(query_embed),
            "vector_preview": str(query_embed[:5]) + " ... [생략]"
        }
    except Exception as e:
        return None, f"임베딩 변환 오류: {e}", []
    
    # Step 3: ClickHouse 벡터 고속 검색 (철저한 사용자 ID 기반 격리 조회)
    db_start = time.time()
    client = get_db_client()
    
    # ⚠️ 보안 필수 규칙: 현재 로그인한 사용자의(user_id) 데이터 조각만 가져올 수 있도록 강제 제한함
    search_query = """
        SELECT content, file_name, file_path, cosineDistance(embedding, %(q)s) as distance, drive_name
        FROM doc_chunks 
        WHERE user_id = %(u)s
        ORDER BY distance ASC 
        LIMIT 4
    """
    
    try:
        result = client.query(search_query, parameters={'q': query_embed, 'u': current_user})
        db_elapsed = time.time() - db_start
        
        retrieved_docs = []
        sources = set()
        context = ""
        
        for row in result.result_rows:
            similarity_pct = (1.0 - float(row[3])) * 100  # 유사도 거리 점수를 % 변환
            retrieved_docs.append({
                "file_name": row[1],
                "file_path": row[2],
                "similarity": f"{similarity_pct:.2f}%",
                "distance": f"{row[3]:.4f}",
                "drive": row[4],
                "content_preview": row[0][:200] + "..."
            })
            context += f"[참고문서: {row[1]} (경로: {row[2]})]\n{row[0]}\n\n"
            sources.add(f"{row[1]} ({row[2]})")
            
        trace_steps['step_3_search'] = {
            "status": f"성공 ({db_elapsed:.3f}초 소요)",
            "total_candidates_found": len(result.result_rows),
            "matched_chunks": retrieved_docs
        }
    except Exception as e:
        return None, f"ClickHouse 지식 검색 오류: {e}", []
    
    if not context:
        trace_steps['step_4_prompt'] = {"status": "미실행", "reason": "검색 결과가 없어 프롬프트를 구성하지 못함"}
        return trace_steps, "죄송합니다. 현재 사용자 권한으로 색인된 클라우드독 파일들 중에서 관련 답변을 이끌어 낼 수 있는 유사 문서를 찾지 못했습니다.", []

    # Step 4: 시스템 프롬프트 합성 및 프롬프트 주입 형태 기록
    system_prompt = f"""당신은 사내 문서 정보를 바탕으로 완벽히 답변하는 스마트 인공지능 어시스턴트입니다. 
제시된 [참고문서] 내용만을 철저하게 근거로 삼아 정직하게 답변해 주세요. 
절대로 주관적이거나 문서에 서술되지 않은 내용을 지어내서 답하면 안 됩니다.
자연스럽게 참고 파일명(출처)을 답변 안에서 언급해 주어 신뢰성을 강화해 주세요.

[참고문서]
{context}
"""
    trace_steps['step_4_prompt'] = {
        "status": "구성 완료",
        "system_prompt_length": len(system_prompt),
        "injected_context_preview": context[:300] + "\n... [중략] ...",
        "full_assembled_prompt": f"[시스템 지침 및 검색 문서 전달 완료] -> 사용자 질의: {query}"
    }

    # Step 5: Ollama gpt-oss 로컬 답변 추론 과정
    try:
        llm_start = time.time()
        response = ollama.generate(model=LLM_MODEL, prompt=f"{system_prompt}\n\n질문: {query}")
        llm_elapsed = time.time() - llm_start
        
        trace_steps['step_5_generation'] = {
            "status": f"성공 ({llm_elapsed:.2f}초 소요)",
            "model_used": LLM_MODEL,
            "total_elapsed_time": f"{time.time() - start_time:.2f}초",
            "tokens_generated": len(response['response'])  # 글자 수로 대략적인 크기 측정
        }
        return trace_steps, response['response'], list(sources)
    except Exception as e:
        return trace_steps, f"Ollama LLM({LLM_MODEL}) 추론 장애 발생: {e}", list(sources)


# --- 8. Streamlit 멀티유저 대응 웹 UI 화면 구성 ---
st.set_page_config(page_title="공용 클라우드독 AI 지식인", layout="wide")
init_clickhouse()

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
    st.session_state.drive_mode = "local"

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
                            
                            local_drives = detect_local_drives()
                            if local_drives:
                                st.session_state.target_drives = local_drives
                                st.session_state.drive_mode = "local"
                            else:
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
        if st.button("🔄 내 전용 변경문서 스마트 동기화", use_container_width=True):
            if st.session_state.target_drives:
                with st.spinner("변경된 대상 분석 중..."):
                    sync_drives_incremental()
            else:
                st.error("동기화할 드라이브 정보가 존재하지 않습니다.")
                
        st.caption("권한이 활성화된 대상 중 새로 생성되거나 변경된 파일 내용만 선별하여 내 계정(ID) 전용으로 저장합니다.")

    # 메인 채팅 기록 표시
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
            # 저장되어 있는 추론 사고 경로 데이터가 있다면 이전 메시지 하단에도 표시해 줍니다.
            if "trace" in msg and msg["trace"]:
                with st.expander("🔍 이 답변을 도출하기 위한 AI 탐색 경로 기록 확인"):
                    st.json(msg["trace"])

    # 질의 입력 및 RAG 추론 실행
    if user_query := st.chat_input("공유 디스크에 대해 질문하세요 (예: VW 인프라 구성도를 보고 싶어요 / 경주공장 대역 IP 대역은?)"):
        st.session_state.messages.append({"role": "user", "content": user_query})
        with st.chat_message("user"):
            st.markdown(user_query)
            
        with st.chat_message("assistant"):
            # Streamlit의 status 객체를 사용하여 단계별로 AI의 추론 흐름을 보여줌
            with st.status("🧠 AI 사고 흐름 및 유사도 검색 엔진 가동 중...", expanded=True) as status_box:
                
                st.write("1️⃣ **사용자 질문 분석 단계**")
                time.sleep(0.3)
                
                st.write("2️⃣ **임베딩 변환 및 질문 벡터화 진행 중...**")
                # 실제 함수 호출 실행
                trace, answer, source_files = ask_question_with_trace(user_query)
                
                if trace:
                    st.write("3️⃣ **ClickHouse 벡터 DB 격리 검색 완수**")
                    st.write("4️⃣ **보안에 입각한 지식 프롬프트 구성 완료**")
                    st.write("5️⃣ **Ollama 로컬 LLM 추론 완료**")
                    status_box.update(label="✅ 추론 및 문서 검색 성공!", state="complete")
                else:
                    status_box.update(label="❌ 탐색 실패", state="error")
            
            # 최종 정돈된 답변 출력
            if source_files:
                answer_with_sources = answer + f"\n\n---\n📌 **연관 출처 파일 리스트:**\n" + "\n".join([f"- {src}" for src in source_files])
            else:
                answer_with_sources = answer

            # 시각적인 Expander로 구체적인 탐색 과정 레이아웃 렌더링
            if trace:
                with st.expander("🔍 AI의 사고 및 문서 검색 과정 (RAG Trace) 상세보기", expanded=False):
                    st.markdown("### 1. 질문 의도 분석")
                    st.json(trace['step_1_intent'])
                    
                    st.markdown("### 2. 질문 벡터화 수치")
                    st.json(trace['step_2_embed'])
                    
                    st.markdown("### 3. ClickHouse DB 매칭 결과 (유사도 역순)")
                    if 'matched_chunks' in trace['step_3_search']:
                        for idx, chunk in enumerate(trace['step_3_search']['matched_chunks']):
                            st.info(
                                f"**🎯 Match {idx+1} [유사도: {chunk['similarity']}]**\n\n"
                                f"* **파일명**: {chunk['file_name']}\n"
                                f"* **경로**: {chunk['file_path']}\n"
                                f"* **본문 요약**: {chunk['content_preview']}"
                            )
                    else:
                        st.write("검색 결과가 없습니다.")
                        
                    st.markdown("### 4. 프롬프트 투입 요약")
                    st.json(trace['step_4_prompt'])
                    
                    st.markdown("### 5. LLM 시간 소모 및 성능 통계")
                    st.json(trace['step_5_generation'])

            st.markdown(answer_with_sources)
            st.session_state.messages.append({
                "role": "assistant", 
                "content": answer_with_sources,
                "trace": trace
            })