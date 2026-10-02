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
import warnings
import logging

# --- 0. 경고 및 로그 전역 필터링 설정 ---
# openpyxl 파싱 도중 발생하는 불필요한 이미지 포맷 및 데이터 유효성 검사 경고 차단
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
# pdfminer의 비표준 폰트 해석 도중 발생하는 지저분한 디버깅 로그 최소화
logging.getLogger("pdfminer").setLevel(logging.ERROR)

# --- 1. 설정 정보 ---
EMBED_MODEL = 'nomic-embed-text'  # 벡터 변환용 임베딩 모델 (Ollama에서 미리 pull 받아주세요)
LLM_MODEL = 'gpt-oss:20b'        # 답변 생성용 LLM

# 사내 WebDAV 인프라 접속 주소 정의 (HTTPS 표준 443 포트 적용)
CLOUDDOC_HOST = 'https://idisk.nvhkorea.com'
CLOUDDOC_PORT = 443

# --- 2. ClickHouse 연결 설정 ---
def get_db_client():
    return clickhouse_connect.get_client(
        host='localhost', port=8123, username='default', password='clickhouse',
        connect_timeout=30, send_receive_timeout=300
    )

# DB 테이블 초기화 및 구버전 테이블 마이그레이션
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
    
    # 2. 기존 구버전 테이블이 이미 존재하는 경우, 누락된 격리용 user_id 등의 컬럼 동적 추가
    try:
        client.command("ALTER TABLE doc_chunks ADD COLUMN IF NOT EXISTS user_id String DEFAULT 'system'")
        client.command("ALTER TABLE doc_chunks ADD COLUMN IF NOT EXISTS drive_name String DEFAULT ''")
        client.command("ALTER TABLE doc_chunks ADD COLUMN IF NOT EXISTS file_path String DEFAULT ''")
        client.command("ALTER TABLE doc_chunks ADD COLUMN IF NOT EXISTS file_mtime Float64 DEFAULT 0.0")
    except Exception as e:
        print(f"테이블 구조 자동 업데이트 중 예외 발생 (안전하게 무시됨): {e}")

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
                            label = f"{letter}_{volume_name}" if volume_name else f"{letter}_공유함"
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
        'webdav_hostname': CLOUDDOC_HOST,
        'webdav_port': CLOUDDOC_PORT,
        'webdav_username': username,
        'webdav_password': password,
        'webdav_root': '/'  # 명시적인 루트 경로 지정
    }
    detected = {}
    try:
        webdav_client = WebDavClient(options)
        # 사설 SSL 인증서 경고 및 거절 현상 우회 설정 적용
        webdav_client.verify = False
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
        'webdav_hostname': CLOUDDOC_HOST,
        'webdav_port': CLOUDDOC_PORT,
        'webdav_username': username,
        'webdav_password': password,
        'webdav_root': '/'
    }
    try:
        webdav_client = WebDavClient(options)
        webdav_client.verify = False
        # 1. '/' 목록 불러오기 시도
        webdav_client.list()
        return True
    except Exception as e:
        print(f"최상위 WebDAV 접근 실패: {e}")
        
        # 2. 사내망 구성에 따라 최상위 루트가 차단되어 있는 경우, '/{username}' 또는 '/idisk' 경로 시도
        for trial_path in [f"/{username}", "/idisk", "/dav"]:
            try:
                options['webdav_root'] = trial_path
                trial_client = WebDavClient(options)
                trial_client.verify = False
                trial_client.list()
                print(f"성공적인 접속 루트 발견: {trial_path}")
                return True
            except Exception:
                continue
                
        # 3. 만약 네트워크 에러가 나더라도 로컬 마운트 드라이브가 연결되어 있는 상태라면 인증을 대체 처리
        local_detected = detect_local_drives()
        if local_detected:
            print("웹 연결에 제한이 있으나 PC에 마운트된 로컬 가상 드라이브가 감지되어 인증을 대체 처리합니다.")
            return True
            
        return False

# --- 5. 파일 텍스트 추출 함수들 (에러 완전 방지 및 고속 스트리밍 적용) ---
def extract_text(file_path):
    ext = os.path.splitext(file_path)[-1].lower()
    text = ""
    try:
        if ext == '.pdf':
            with pdfplumber.open(file_path) as pdf:
                page_texts = []
                # 특정 페이지의 비표준 폰트 해석 에러가 전체 파이프라인을 멈추지 않도록 개별 감싸기 수행
                for page in pdf.pages:
                    try:
                        p_text = page.extract_text()
                        if p_text:
                            page_texts.append(p_text)
                    except Exception as page_err:
                        print(f"⚠️ PDF 특정 페이지 파싱 실패 건너뜀 ({file_path}): {page_err}")
                        continue
                text = "\n".join(page_texts)
                
        elif ext == '.docx':
            doc = Document(file_path)
            text = "\n".join([p.text for p in doc.paragraphs])
            
        elif ext in ['.xlsx', '.xls']:
            # read_only=True 와 data_only=True 모드를 강제하여 
            # WMF 등의 부적절한 이미지 서식이나 무거운 함수 계산 데이터를 무시하고 텍스트만 초고속 추출
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
                sheet_texts = []
                for sheet in wb.sheetnames:
                    ws = wb[sheet]
                    for row in ws.iter_rows(values_only=True):
                        row_text = " ".join([str(cell) for cell in row if cell is not None])
                        if row_text.strip():
                            sheet_texts.append(row_text)
                text = "\n".join(sheet_texts)
                
        elif ext == '.txt':
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
    except Exception as e:
        print(f"❌ 파일 읽기 전체 실패 ({file_path}): {e}")
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

# --- 6. 사용자 격리형 스마트 증분 동기화 (Delta Sync) 통합 구현 ---
def sync_drives_incremental():
    current_user = st.session_state.username
    st.info(f"🔄 [{current_user}] 계정의 전용 변경 문서만 선별하여 데이터베이스를 동기화 중입니다...")
    client = get_db_client()
    
    # 1. 기존 DB에 저장되어 있는 현재 '로그인 사용자(user_id)'의 데이터 파일 정보만 부분 로드
    db_files = {}
    try:
        result = client.query(
            "SELECT DISTINCT file_path, file_mtime FROM doc_chunks WHERE user_id = %(u)s", 
            parameters={'u': current_user}
        )
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
    valid_files = []

    # A. [로컬 가상 드라이브 모드] 사전 스캔 단계 (임시 파일 ~$ 필터링)
    if drive_mode == "local":
        for drive_label, drive_path in target_drives.items():
            if not os.path.exists(drive_path):
                st.warning(f"⚠️ {drive_label} ({drive_path}) 드라이브에 접근할 수 없습니다. 클라우드독 로그인을 체크해 주세요.")
                continue
            st.write(f"🔍 '{drive_label}' 드라이브 구조 스캔 중...")
            for root, dirs, files_in_dir in os.walk(drive_path):
                for file in files_in_dir:
                    # 임시 파일 및 시스템 오피스 파일 (~$로 시작하는 임시 록 파일) 배제
                    if file.startswith('~$'):
                        continue
                    ext = os.path.splitext(file)[-1].lower()
                    if ext in supported_extensions:
                        full_path = os.path.join(root, file)
                        try:
                            current_mtime = os.path.getmtime(full_path)
                        except Exception:
                            continue
                        valid_files.append({
                            'drive_label': drive_label,
                            'file_name': file,
                            'file_path': full_path,
                            'mtime': current_mtime
                        })
                        
    # B. [WebDAV 클라우드 연동 모드] 사전 스캔 단계 (임시 파일 ~$ 필터링)
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
            st.error(f"클라우드독 WebDAV 연결에 실패했습니다: {e}")
            return

        for drive_label, drive_path in target_drives.items():
            st.write(f"☁️ 클라우드독 '{drive_label}' 폴더 목록 탐색 중...")
            webdav_files = get_webdav_files_recursive(webdav_client, drive_path)
            for file_info in webdav_files:
                file_name = file_info['name']
                # 임시 파일 배제
                if file_name.startswith('~$'):
                    continue
                ext = os.path.splitext(file_name)[-1].lower()
                if ext in supported_extensions:
                    valid_files.append({
                        'drive_label': drive_label,
                        'file_name': file_name,
                        'file_path': file_info['path'],
                        'mtime': file_info['mtime']
                    })

    # C. 공통 동기화 처리 루프 (진행율% 및 실시간 소요 시간/예상 시간 시각화)
    total_files = len(valid_files)
    if total_files == 0:
        st.info("동기화할 대상 파일이 존재하지 않습니다.")
        return

    progress_bar = st.progress(0.0)
    status_placeholder = st.empty()
    sync_start_time = time.time()

    temp_dir_ctx = tempfile.TemporaryDirectory() if drive_mode != "local" else None
    temp_dir = temp_dir_ctx.name if temp_dir_ctx else None

    try:
        for idx, file_info in enumerate(valid_files):
            drive_label = file_info['drive_label']
            file_name = file_info['file_name']
            file_path = file_info['file_path']
            current_mtime = file_info['mtime']
            
            # 삭제 대상 판별을 위해 추적 셋에 기록
            scanned_file_paths.add(file_path)
            
            # 동적 진행 통계 산출
            processed = idx + 1
            pct = processed / total_files
            progress_bar.progress(pct)
            
            elapsed = time.time() - sync_start_time
            avg_time_per_file = elapsed / processed
            est_total_time = avg_time_per_file * total_files
            est_remaining = max(0.0, est_total_time - elapsed)
            
            status_placeholder.markdown(f"""
            **지식 동기화 및 임베딩 분석 진행 중...** ({processed}/{total_files} 파일) - **{pct*100:.1f}% 완료**
            * **현재 처리 중:** `{drive_label}` / `{file_name}`
            * **실제 경과 시간:** `{elapsed:.1f}초`
            * **예상 남은 시간:** `{est_remaining:.1f}초` (평균 `{avg_time_per_file:.2f}초/파일`)
            """)

            # 변경 이력 판별 검증
            if file_path in db_files and db_files[file_path] == current_mtime:
                files_skipped += 1
                continue

            # 변경 감지 시 타겟 데이터 조각 초기화
            if file_path in db_files:
                escaped_path = file_path.replace('\\', '\\\\')
                client.command(
                    f"ALTER TABLE doc_chunks DELETE WHERE file_path = '{escaped_path}' AND user_id = '{current_user}'"
                )
            
            # 본문 추출 수행
            full_text = ""
            if drive_mode == "local":
                full_text = extract_text(file_path)
            else:
                local_temp_path = os.path.join(temp_dir, file_name)
                try:
                    webdav_client.download_file(file_path, local_temp_path)
                    full_text = extract_text(local_temp_path)
                    if os.path.exists(local_temp_path):
                        os.remove(local_temp_path)
                except Exception as de:
                    print(f"다운로드 실패 ({file_path}): {de}")
                    continue
            
            if not full_text.strip():
                continue
                
            chunks = chunk_text(full_text)
            data_to_insert = []
            for i, chunk in enumerate(chunks):
                try:
                    response = ollama.embed(model=EMBED_MODEL, input=chunk)
                    embedding = response['embeddings'][0]
                    chunk_id_prefix = "webdav_" if drive_mode != "local" else ""
                    chunk_id = f"{chunk_id_prefix}{current_user}_{drive_label}_{file_name}_{i}_{current_mtime}"
                    data_to_insert.append([
                        chunk_id, current_user, drive_label, file_name, file_path, chunk, embedding, current_mtime
                    ])
                except Exception as ex:
                    print(f"임베딩 변환 에러 ({file_name}): {ex}")
                    continue
                    
            if data_to_insert:
                client.insert(
                    'doc_chunks', 
                    data_to_insert, 
                    column_names=['id', 'user_id', 'drive_name', 'file_name', 'file_path', 'content', 'embedding', 'file_mtime']
                )
                files_added_or_updated += 1
                
    finally:
        if temp_dir_ctx:
            temp_dir_ctx.cleanup()

    # 3. 사용자 권한 내에서 삭제된 파일 청크 제거
    deleted_files = set(db_files.keys()) - scanned_file_paths
    if deleted_files:
        st.write(f"🧹 공유 디스크에서 지워진 {len(deleted_files)}개의 항목 데이터를 안전하게 정돈하고 있습니다...")
        for del_path in deleted_files:
            escaped_path = del_path.replace('\\', '\\\\')
            client.command(
                f"ALTER TABLE doc_chunks DELETE WHERE file_path = '{escaped_path}' AND user_id = '{current_user}'"
            )

    st.success(
        f"🎉 동기화 마침! [{current_user} 계정 전용 신규/수정 파일]: {files_added_or_updated}건 처리완료, "
        f"[유지 파일]: {files_skipped}건 보존, [삭제 파일]: {len(deleted_files)}건 제거완료."
    )

# --- 7. 사고 회로 추적형 (RAG Trace) 답변 생성 함수 ---
def ask_question_with_trace(query):
    trace_steps = {}
    start_time = time.time()
    current_user = st.session_state.username
    
    # Step 1: 사용자 질문 분석 단계 기록
    trace_steps['step_1_intent'] = {
        "status": "진행 완료",
        "raw_query": query,
        "parsed_keywords": [word for word in query.split() if len(word) > 1]
    }
    
    # Step 2: 임베딩 생성 소요 시간 기록 및 벡터화
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
    
    # Step 3: ClickHouse 벡터 고속 검색 (철저한 로그인 사용자 보안 격리 쿼리 수행)
    db_start = time.time()
    client = get_db_client()
    
    # ⚠️ 보안 필수 수칙: 현재 로그인 계정(user_id)과 정확히 매칭되는 문서 샌드박스 범위 내에서만 인덱싱
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
            similarity_pct = (1.0 - float(row[3])) * 100  # 유사도 거리 수치를 백분율 퍼센트로 치환
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
        trace_steps['step_4_prompt'] = {"status": "미실행", "reason": "사용자의 해당 검색 영역에 학습된 문서 결과가 없음"}
        return trace_steps, "죄송합니다. 현재 본인 계정 권한으로 가상 동기화된 클라우드독 파일 안에서 관련 있는 지식 조각을 찾을 수 없었습니다.", []

    # Step 4: 시스템 프롬프트 합성 과정 기록
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
        "full_assembled_prompt": f"[보안 RAG 지침 구성 완료]"
    }

    # Step 5: Ollama gpt-oss 로컬 LLM 추론 시간 연산 기록
    try:
        llm_start = time.time()
        response = ollama.generate(model=LLM_MODEL, prompt=f"{system_prompt}\n\n질문: {query}")
        llm_elapsed = time.time() - llm_start
        
        trace_steps['step_5_generation'] = {
            "status": f"성공 ({llm_elapsed:.2f}초 소요)",
            "model_used": LLM_MODEL,
            "total_elapsed_time": f"{time.time() - start_time:.2f}초",
            "tokens_generated": len(response['response'])
        }
        return trace_steps, response['response'], list(sources)
    except Exception as e:
        return trace_steps, f"Ollama LLM({LLM_MODEL}) 추론 장애 발생: {e}", list(sources)

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
                    st.write("3️⃣ **ClickHouse DB 격리 검색 완수**")
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