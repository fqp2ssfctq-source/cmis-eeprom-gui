"""
EEPROM CMIS Writer  v2.9.4
CP2112-F03-GM USB-HID I2C 브릿지 기반 CMIS 5.0 EEPROM 관리 도구

구조:
  CP2112I2C    - Silicon Labs CP2112 USB-HID → I2C 통신 드라이버
  SimulatedI2C - 하드웨어 없이 동작 확인용 시뮬레이터
  CanvasTable  - Canvas 기반 커스텀 테이블 위젯
                 (열 너비 드래그 조절 / 셀 선택 복사 / 바이트 팝업 편집)
  App          - 메인 GUI 애플리케이션 (tkinter)

탭 구성:
  연결   - CP2112 장치 선택 및 I2C 연결
  EEPROM 뷰어 - 페이지별 CMIS 메모리 맵 표시 및 편집
  쓰기   - EEPROM 실제 쓰기 및 검증
  로그   - 동작 이력
  정보   - 버전 및 시스템 정보
"""
# ── 표준 라이브러리 및 GUI 프레임워크 ────────────────────
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading, time, os, sys, copy, colorsys, traceback, logging

# ── 앱 메타 정보 ──────────────────────────────────────────
# TYPE_BADGE/TYPE_COLOR: 버전 히스토리 표시용 배지
#   [M]=major(하위 호환 불가), [+]=minor(기능 추가), [·]=patch(버그 수정)
APP_NAME    = "EEPROM CMIS Writer"
APP_VERSION = "3.1.2"
APP_DATE    = "2026-04-06"
CHANGELOG = [
    ("3.0.5","2026-04-07","patch",  "lambda 클로저 변수 캡처 오류 수정"),
    ("3.0.6","2026-04-07","major",  "USB 캡처 분석 기반: 실제 동작과 동일한 시퀀스로 교체"),
    ("3.0.6","2026-04-07","major",  "패킷 분석: 8bit 주소 사용 / 128바이트 일괄 읽기로 교체"),
    ("3.1.2","2026-04-07","minor",  "P01h/P02h 핵심 필드 Decoded/Description 추가"),
    ("3.1.1","2026-04-07","patch",  "바이트팝업 인덱스 dec 주소 표시 / I2C 기본주소 0x50 수정"),
    ("3.1.0","2026-04-07","minor",  "I2C 기본주소 0x50 / 읽기 진행바 / 작업 시에만 포트 점유"),
    ("3.0.7","2026-04-07","patch",  "read_page: ForceReadResponse 추가, 읽기 시퀀스 LabVIEW VI와 동일하게"),
    ("3.0.6","2026-04-07","patch",  "버그수정: read_bytes→read_page, 8bit주소 변환, SimulatedI2C read_page 추가"),
    ("3.0.5","2026-04-07","minor",  "I2C 주소 스캔 기능 추가, 스캔 결과로 주소 자동 설정"),
    ("3.0.4","2026-04-07","patch",  "전체 DLL 함수 시그니처 헤더파일 기준으로 교정"),
    ("3.0.3","2026-04-07","patch",  "HidSmbus_Open 파라미터 순서 수정 (deviceNum, &handle)"),
    ("3.0.2","2026-04-07","patch",  "장치 이름 GetAttributes로 VID/PID/S/N 표시"),
    ("3.0.1","2026-04-07","patch",  "장치 이름 문자열 깨짐 수정 (unicode→bytes 버퍼)"),
    ("3.0.0","2026-04-07","major",  "hidapi raw HID -> SLABHIDtoSMBus.dll ctypes 방식으로 교체"),
    ("2.9.6","2026-04-07","patch",  "crash.log 인코딩 수정 / OSError: read error 처리 추가"),
    ("2.9.5","2026-04-07","patch",  "CP2112 장치 열림 상태 추적 / ValueError→IOError 변환"),
    ("2.9.4","2026-04-06","patch",  "코드 전체 주석 추가"),
    ("2.9.3","2026-04-06","patch",  "바이트별 개별 Entry 팝업으로 편집"),
    ("2.9.2","2026-04-06","patch",  "멀티바이트 Value 편집 시 전체 바이트 표시 및 수정"),
    ("2.9.1","2026-04-06","patch",  "Decoded 미해석 필드 N/A 표기, 편집 시 Decoded 동시 갱신"),
    ("2.9.0","2026-04-06","minor",  "_safe_call 제거 / Value[hex] 인라인 직접 편집"),
    ("2.8.1","2026-04-06","patch",  "crash.log: 실행 중 모든 예외 기록"),
    ("2.8.0","2026-04-06","minor",  "코드 정리: 중복/미사용 코드 제거"),
    ("2.7.4","2026-04-06","patch",  "hidapi 없을 때 자동 설치"),
    ("2.7.3","2026-04-06","patch",  "I2C 주소 0x50 고정"),
    ("2.7.2","2026-04-06","patch",  "ver_label 색상 덮어쓰기 방지"),
    ("2.7.1","2026-04-06","patch",  "HID_OK 참조 수정"),
    ("2.7.0","2026-04-06","minor",  "CP2112 USB-HID I2C 브릿지 지원"),
    ("2.6.5","2026-04-06","minor",  "셀/범위 드래그 선택 + Ctrl+C 복사"),
    ("2.6.4","2026-04-06","patch",  "기본 모드: 라이트 100%"),
    ("2.6.3","2026-04-06","patch",  "Value[hex] 전체 바이트 표시"),
    ("2.6.2","2026-04-06","patch",  "create_text width= 제거"),
    ("2.6.1","2026-04-06","minor",  "헤더 드래그 열 너비 조절"),
    ("2.6.0","2026-04-06","major",  "CanvasTable 교체: 열 너비 완전 제어"),
    ("2.5.0","2026-04-06","minor",  "P00h Value 해석 강화 / Decoded 열"),
    ("2.4.0","2026-04-06","minor",  "6컬럼 구성 / VendorSN DateCode 수정"),
    ("2.3.0","2026-04-06","minor",  "MultiColumn Treeview / 열 구분선"),
    ("2.2.0","2026-04-06","patch",  "Canvas 오버레이 제거"),
    ("2.1.0","2026-04-06","minor",  "열 간격 / 라이트모드 밝기 유지"),
    ("2.0.0","2026-04-06","major",  "버전 관리 / 능동 밝기 조절"),
    ("1.1.0","2026-04-05","minor",  "시뮬레이션 모드 추가"),
    ("1.0.0","2026-04-04","major",  "최초 릴리즈"),
]
TYPE_BADGE = {"major":"[M]","minor":"[+]","patch":"[·]"}
TYPE_COLOR = {"major":"#f07070","minor":"#7ab3f5","patch":"#888888"}

# ── crash.log 설정 ────────────────────────────────────────
def _setup_crash_log():
    """실행 파일과 같은 폴더에 crash.log 생성, 모든 예외 기록"""
    log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "crash.log")
    # FileHandler로 직접 생성 → Windows 한글 환경에서 인코딩 명시
    handler=logging.FileHandler(log_path, encoding="utf-8", mode="a")
    handler.setLevel(logging.ERROR)
    handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    ))
    root_logger=logging.getLogger()
    root_logger.setLevel(logging.ERROR)
    root_logger.addHandler(handler)
    # 미처리 예외 자동 기록
    def _exc_handler(exc_type, exc_value, exc_tb):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_tb)
            return
        logging.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_tb))
        # 원래 동작도 유지
        sys.__excepthook__(exc_type, exc_value, exc_tb)
    sys.excepthook = _exc_handler
    return log_path

CRASH_LOG = _setup_crash_log()

# ── SLABHIDtoSMBus.dll 로드 ──────────────────────────────
# Silicon Labs 전용 고수준 API (LabVIEW VI와 동일한 방식)
import ctypes

def _load_slab_dll():
    """SLABHIDtoSMBus.dll 로드. 스크립트 폴더 우선 탐색."""
    if sys.platform != "win32":
        return False, None
    dirs = [os.path.dirname(os.path.abspath(__file__)), os.getcwd()]
    for d in dirs:
        path = os.path.join(d, "SLABHIDtoSMBus.dll")
        if os.path.exists(path):
            try:
                return True, ctypes.WinDLL(path)
            except Exception as e:
                print(f"DLL 로드 실패: {e}")
    print("SLABHIDtoSMBus.dll 없음. 스크립트와 같은 폴더에 복사하세요.")
    return False, None

HID_OK, _SLAB_DLL = _load_slab_dll()
CP2112_VID = 0x10C4
CP2112_PID = 0xEA90

# ── 밝기 유틸 ──────────────────────────────────────────────
# _h2rgb : "#RRGGBB" → (r,g,b) 0~1 float 변환
# _rgb2h : (r,g,b) → "#RRGGBB" 변환
# _adj   : hex 색상을 HSV 공간에서 밝기(f) 조정
# make_theme: 베이스 테마 dict를 밝기 factor로 조정한 새 dict 반환
#             SKIP 키(강조색 등)는 밝기 조정 제외
def _h2rgb(h):
    h=h.lstrip("#")
    return tuple(int(h[i:i+2],16)/255 for i in(0,2,4))
def _rgb2h(r,g,b):
    return "#{:02X}{:02X}{:02X}".format(int(r*255),int(g*255),int(b*255))
def _adj(hc,f):
    r,g,b=_h2rgb(hc); hh,s,v=colorsys.rgb_to_hsv(r,g,b)
    return _rgb2h(*colorsys.hsv_to_rgb(hh,s,max(0.0,min(1.0,v*f))))
SKIP={"acc","grn","amb","red","dirty","sel_fg","entry_fg"}  # 밝기 조정 제외 키
def make_theme(base,f):
    return {k:(_adj(v,f) if isinstance(v,str) and v.startswith("#") and k not in SKIP else v)
            for k,v in base.items()}

# ── 테마 ──────────────────────────────────────────────────
# 다크/라이트 베이스 테마 dict
# bg0~bg3 : 배경 계층 (bg0=가장 어두움, bg3=가장 어두운 강조)
# t1~t4   : 텍스트 계층 (t1=주 텍스트, t4=비활성)
# acc      : 강조색 (탭 활성, 헤더 텍스트 등)
# col0_*   : Addr 열 배경 (짝수/홀수 행)
# col1_*   : Data 열 배경 (Value, Decoded, Field, Description)
# col_sep  : 열 구분선 색상
_DARK={
    "bg0":"#1a1a1a","bg1":"#252525","bg2":"#202020","bg3":"#161616",
    "bd":"#3a3a3a","bd2":"#505050",
    "t1":"#e0e0e0","t2":"#aaaaaa","t3":"#686868","t4":"#444444",
    "acc":"#7ab3f5","grn":"#6ac87a","amb":"#f0c060","red":"#f07070","dirty":"#f0a050",
    "sel_bg":"#1e3a5f","sel_fg":"#7ab3f5",
    "entry_bg":"#161616","entry_fg":"#f0c060",
    "btn_bg":"#2a2a2a","btn_fg":"#d0d0d0",
    # 열별 배경 (짝수열 / 홀수열)
    "col0_even":"#1c1c1c","col0_odd":"#232323",
    "col1_even":"#1c1c1c","col1_odd":"#232323",
    "col2_even":"#1c1c1c","col2_odd":"#232323",
    "col3_even":"#1c1c1c","col3_odd":"#232323",
    "col4_even":"#1c1c1c","col4_odd":"#232323",
    "col_head":"#2e2e2e","col_sep":"#404040",
    "row_hl":"#2a2a2a",
}
_LIGHT={
    "bg0":"#d2d2d2","bg1":"#c6c6c6","bg2":"#cccccc","bg3":"#c0c0c0",
    "bd":"#9a9a9a","bd2":"#787878",
    "t1":"#111111","t2":"#333333","t3":"#666666","t4":"#999999",
    "acc":"#185FA5","grn":"#2a5a0a","amb":"#7a4a00","red":"#8a2020","dirty":"#7a4000",
    "sel_bg":"#b8d4ee","sel_fg":"#0a3a70",
    "entry_bg":"#bbbbbb","entry_fg":"#7a4a00",
    "btn_bg":"#c8c8c8","btn_fg":"#111111",
    "col0_even":"#c0c0c0","col0_odd":"#cccccc",
    "col1_even":"#c0c0c0","col1_odd":"#cccccc",
    "col2_even":"#c0c0c0","col2_odd":"#cccccc",
    "col3_even":"#c0c0c0","col3_odd":"#cccccc",
    "col4_even":"#c0c0c0","col4_odd":"#cccccc",
    "col_head":"#b4b4b4","col_sep":"#909090",
    "row_hl":"#e8e8e8",
}

# ── CMIS 상수 ─────────────────────────────────────────────
# SFF8024 : SFF-8024 모듈 타입 식별자 테이블
# I2C_ADDR_A0 : CMIS 모듈 I2C 8bit 주소 (7bit=0x50)
# PAGE_KEYS   : 내부 데이터 딕셔너리 키 (소문자)
# PAGE_LABELS : GUI 표시용 레이블 (대문자)
# COL_CFG     : 테이블 열 정의
#   (헤더문자열, 기본폭px, 최소폭px, col_key, stretch여부)
#   col_key=0 → Addr 배경, col_key=1 → Data 배경
SFF8024={0x18:"QSFP-DD",0x19:"OSFP-8X",0x1A:"SFP-DD",
         0x11:"QSFP28",0x0D:"QSFP+",0x03:"SFP",0x0C:"QSFP"}
I2C_ADDR_A0=0xA0
PAGE_LABELS=["A0h","P00h","P01h","P02h","P10h","P11h"]
PAGE_KEYS  =["a0","p00","p01","p02","p10","p11"]

# 열 정의: (헤더, 기본폭, 최소폭, col_key, stretch)
COL_CFG=[
    ("Addr(dec)",  80,  55, 0, False),   # Addr: 밝은 배경
    ("Addr(hex)",  90,  65, 0, False),   # Addr: 밝은 배경
    ("Value[hex]",140,  90, 1, False),   # Data: 어두운 배경 (요약 표시)
    ("Decoded",   220, 140, 1, False),   # Data: 어두운 배경 (전체 내용)
    ("Field",     185, 115, 1, False),   # Data: 어두운 배경
    ("Description",320,140, 1, True),   # Data: 어두운 배경
]

# h2         : 정수 → 2자리 대문자 16진수 문자열 (None이면 "--")
# asc_display: 바이트 배열 → ASCII 표시 문자열
#              0x00=·, 0x20~0x7E=그대로, 나머지=?
# parse_txt  : 탭 구분 txt 파일 → 페이지별 바이트 배열 dict
# build_txt  : 페이지별 바이트 배열 dict → txt 파일 문자열
def h2(v): return "--" if v is None else f"{int(v):02X}"

def i2c_8bit(addr_var_str):
    """i2c_addr_var 문자열(7bit hex) → 8bit 주소 정수
    예: "50" → 0xA0, "51" → 0xA2
    """
    return (int(addr_var_str.strip(), 16) & 0x7F) << 1
def asc_display(a,s,e):
    """표시용: 유효문자 + 공백은 그대로, 0x00은 · 로 표시"""
    r=""
    for i in range(s,e+1):
        c=a[i] if i<len(a) else 0
        if c==0x00: r+="·"
        elif 0x20<=c<=0x7E: r+=chr(c)
        else: r+="?"
    return r.rstrip("·").rstrip() or "(empty)"
def parse_txt(text):
    data={k:[0]*128 for k in PAGE_KEYS}
    for line in text.strip().splitlines():
        p=line.strip().split()
        if not p or p[0].lower()=="addr": continue
        try: a=int(p[0],16)
        except: continue
        if len(p)>1 and p[1]!="--": data["a0"][a]=int(p[1],16)
        if len(p)>2:
            try: idx=int(p[2],16)-0x80
            except: continue
            for ki,pk in enumerate(["p00","p01","p02","p10","p11"]):
                col=3+ki
                if col<len(p) and p[col]!="--": data[pk][idx]=int(p[col],16)
    return data
def build_txt(data):
    lines=["Addr\tA0\tAddr\tP00h\tP01h\tP02h\tP10h\tP11h"]
    for i in range(128):
        row=[h2(i),h2(data["a0"][i]),h2(i+0x80)]
        for pk in["p00","p01","p02","p10","p11"]: row.append(h2(data[pk][i]))
        lines.append("\t".join(row))
    return "\n".join(lines)

# ── CP2112 USB-HID I2C (SLABHIDtoSMBus.dll ctypes) ───────
# 패킷 분석 결과 기반:
#   - slaveAddress: 8bit 주소 그대로 전달 (0xA0)
#   - 128바이트 일괄 읽기 (AddressReadRequest 1회)
#   - GetReadResponse를 반복 호출하여 61바이트씩 수신
class CP2112I2C:
    """CP2112-F03-GM SLABHIDtoSMBus.dll ctypes 드라이버"""
    HID_SMBUS_SUCCESS = 0x00
    XFER_IDLE     = 0x00
    XFER_BUSY     = 0x01
    XFER_COMPLETE = 0x02
    XFER_ERROR    = 0x03

    def __init__(self, device_index=0):
        if not HID_OK or _SLAB_DLL is None:
            raise RuntimeError("SLABHIDtoSMBus.dll 없음. 스크립트와 같은 폴더에 복사하세요.")
        self._dll = _SLAB_DLL
        self._handle = ctypes.c_void_p(0)
        self._is_open = False
        # HidSmbus_Open(HID_SMBUS_DEVICE* device, DWORD deviceNum, WORD vid, WORD pid)
        status = self._dll.HidSmbus_Open(
            ctypes.byref(self._handle),
            ctypes.c_uint32(device_index),
            ctypes.c_uint16(CP2112_VID),
            ctypes.c_uint16(CP2112_PID))
        if status != self.HID_SMBUS_SUCCESS:
            raise IOError(f"HidSmbus_Open 실패 (status=0x{status:02X})")
        self._is_open = True
        self._configure()

    def _configure(self):
        """HidSmbus_SetSmbusConfig: 400kHz, writeTimeout=1000, readTimeout=1000, retries=3"""
        st = self._dll.HidSmbus_SetSmbusConfig(
            self._handle,
            ctypes.c_uint32(400000),
            ctypes.c_uint8(0x02),
            ctypes.c_bool(False),
            ctypes.c_uint16(1000),
            ctypes.c_uint16(1000),
            ctypes.c_bool(True),
            ctypes.c_uint16(3))
        if st != self.HID_SMBUS_SUCCESS:
            raise IOError(f"HidSmbus_SetSmbusConfig 실패 (0x{st:02X})")

    def _check_open(self):
        if not self._is_open:
            raise IOError("CP2112 미연결. 재연결하세요.")

    def _wait_complete(self, timeout_ms=2000):
        """TransferStatusRequest + GetTransferStatusResponse 폴링"""
        self._check_open()
        deadline = time.time() + timeout_ms/1000.0
        ts = ctypes.c_uint8(0)
        ds = ctypes.c_uint8(0)
        nr = ctypes.c_uint16(0)
        br = ctypes.c_uint16(0)
        while time.time() < deadline:
            self._dll.HidSmbus_TransferStatusRequest(self._handle)
            st = self._dll.HidSmbus_GetTransferStatusResponse(
                self._handle,
                ctypes.byref(ts), ctypes.byref(ds),
                ctypes.byref(nr), ctypes.byref(br))
            if st != self.HID_SMBUS_SUCCESS:
                raise IOError(f"GetTransferStatusResponse 실패 (0x{st:02X})")
            v = ts.value
            if v == self.XFER_COMPLETE: return br.value
            if v == self.XFER_ERROR:
                raise IOError(
                    f"I2C 전송 오류 (S1=0x{ds.value:02X}). "
                    f"배선/주소/풀업저항 확인 필요")
            time.sleep(0.002)
        raise IOError(f"CP2112 타임아웃 ({timeout_ms}ms)")

    def close(self):
        self._is_open = False
        try: self._dll.HidSmbus_Close(self._handle)
        except: pass

    def write_byte(self, i2c_addr_8bit, reg, value):
        """1바이트 쓰기. i2c_addr_8bit: 8bit 주소 (예: 0xA0)"""
        self._check_open()
        # WriteRequest: [reg, value] 2바이트 전송
        buf = (ctypes.c_uint8 * 2)(reg & 0xFF, value & 0xFF)
        st = self._dll.HidSmbus_WriteRequest(
            self._handle,
            ctypes.c_uint8(i2c_addr_8bit & 0xFE),  # 8bit write 주소 (bit0=0)
            buf,
            ctypes.c_uint8(2))
        if st != self.HID_SMBUS_SUCCESS:
            raise IOError(f"HidSmbus_WriteRequest 실패 (0x{st:02X})")
        self._wait_complete()
        return True

    def read_page(self, i2c_addr_8bit, start_reg, num_bytes=128):
        """
        여러 바이트 읽기 (LabVIEW VI와 동일한 시퀀스)
        1. HidSmbus_AddressReadRequest  - 읽기 요청
        2. TransferStatusRequest + GetTransferStatusResponse - 완료 대기
        3. HidSmbus_ForceReadResponse   - 읽기 응답 강제 요청
        4. HidSmbus_GetReadResponse     - 61바이트씩 반복 수신
        """
        self._check_open()
        target_addr = (ctypes.c_uint8 * 16)(start_reg & 0xFF)
        st = self._dll.HidSmbus_AddressReadRequest(
            self._handle,
            ctypes.c_uint8(i2c_addr_8bit & 0xFE),
            ctypes.c_uint16(num_bytes),
            ctypes.c_uint8(1),
            target_addr)
        if st != self.HID_SMBUS_SUCCESS:
            raise IOError(f"HidSmbus_AddressReadRequest 실패 (0x{st:02X})")
        # 전송 완료 대기
        self._wait_complete()
        # ForceReadResponse: 장치에게 데이터를 보내달라고 요청
        # LabVIEW GetReadData.vi: TransferStatus 후 ForceReadResponse 호출
        remaining = num_bytes
        result = []
        while remaining > 0:
            chunk = min(remaining, 61)
            st_f = self._dll.HidSmbus_ForceReadResponse(
                self._handle,
                ctypes.c_uint16(chunk))
            if st_f != self.HID_SMBUS_SUCCESS:
                break
            # GetReadResponse로 수신
            rs  = ctypes.c_uint8(0)
            buf = (ctypes.c_uint8 * 61)()
            nr  = ctypes.c_uint8(0)
            st2 = self._dll.HidSmbus_GetReadResponse(
                self._handle,
                ctypes.byref(rs), buf,
                ctypes.c_uint8(61), ctypes.byref(nr))
            if st2 != self.HID_SMBUS_SUCCESS or nr.value == 0:
                break
            result.extend(buf[j] for j in range(nr.value))
            remaining -= nr.value
        # 부족하면 0으로 채움
        while len(result) < num_bytes:
            result.append(0)
        return result[:num_bytes]

    def read_byte(self, i2c_addr_8bit, reg):
        """1바이트 읽기 (호환성 유지)"""
        result = self.read_page(i2c_addr_8bit, reg, 1)
        if not result:
            raise IOError(f"읽기 데이터 없음 reg=0x{reg:02X}")
        return result[0]

    def set_page(self, page_num):
        """CMIS 페이지 전환: Byte 0x7F에 페이지 번호 쓰기"""
        self.write_byte(I2C_ADDR_A0, 0x7F, page_num)
        time.sleep(0.01)

    def scan_addresses(self, start=0x08, end=0x77):
        """I2C 주소 스캔 (7bit 범위 → 8bit로 변환하여 시도)"""
        self._check_open()
        found = []
        for addr_7bit in range(start, end+1):
            addr_8bit = addr_7bit << 1
            try:
                buf = (ctypes.c_uint8 * 1)(0x00)
                st = self._dll.HidSmbus_WriteRequest(
                    self._handle,
                    ctypes.c_uint8(addr_8bit),
                    buf, ctypes.c_uint8(1))
                if st != self.HID_SMBUS_SUCCESS:
                    continue
                ts = ctypes.c_uint8(0)
                ds = ctypes.c_uint8(0)
                nr = ctypes.c_uint16(0)
                br = ctypes.c_uint16(0)
                deadline = time.time() + 0.15
                while time.time() < deadline:
                    self._dll.HidSmbus_TransferStatusRequest(self._handle)
                    self._dll.HidSmbus_GetTransferStatusResponse(
                        self._handle,
                        ctypes.byref(ts), ctypes.byref(ds),
                        ctypes.byref(nr), ctypes.byref(br))
                    if ts.value in (0x02, 0x03, 0x00): break
                    time.sleep(0.002)
                if ts.value == 0x02:
                    found.append(addr_7bit)
            except Exception:
                continue
        return found

    @staticmethod
    def list_devices():
        """연결된 CP2112 장치 목록 반환"""
        if not HID_OK or _SLAB_DLL is None: return []
        num = ctypes.c_uint32(0)
        _SLAB_DLL.HidSmbus_GetNumDevices(
            ctypes.byref(num),
            ctypes.c_uint16(CP2112_VID),
            ctypes.c_uint16(CP2112_PID))
        result = []
        for i in range(num.value):
            vid = ctypes.c_uint16(0)
            pid = ctypes.c_uint16(0)
            rel = ctypes.c_uint16(0)
            _SLAB_DLL.HidSmbus_GetAttributes(
                ctypes.c_uint32(i),
                ctypes.c_uint16(CP2112_VID), ctypes.c_uint16(CP2112_PID),
                ctypes.byref(vid), ctypes.byref(pid), ctypes.byref(rel))
            sn_buf = ctypes.create_string_buffer(260)
            _SLAB_DLL.HidSmbus_GetString(
                ctypes.c_uint32(i),
                ctypes.c_uint16(CP2112_VID), ctypes.c_uint16(CP2112_PID),
                sn_buf, ctypes.c_uint32(0x04))
            sn = sn_buf.value.decode("ascii","ignore").strip()
            prod = f"CP2112 [{i}]"
            if sn: prod += f" S/N:{sn}"
            result.append({"product_string": prod,
                           "serial_number":  sn,
                           "manufacturer_string": "Silicon Labs"})
        return result

class SimulatedI2C:
    """하드웨어 없이 GUI 동작 확인용 시뮬레이터."""
    simulated=True
    def __init__(self): self._mem=[0]*256
    def close(self): pass
    def read_byte(self,*a): time.sleep(0.001); return 0
    def read_page(self,*a,**kw):
        n=kw.get("num_bytes",128) if len(a)<3 else a[2]
        time.sleep(0.001); return [0]*n
    def write_byte(self,*a): time.sleep(0.001); return True
    def set_page(self,*a): pass
    def scan_addresses(self,*a,**kw): return []

# ── CanvasTable ───────────────────────────────────────────
class CanvasTable(tk.Frame):
    """
    Canvas 기반 커스텀 테이블 위젯.
    tkinter 기본 Treeview 대신 사용하는 이유:
      - Treeview는 열별 배경색 독립 제어 불가
      - Treeview는 pack 레이아웃에서 열 너비가 자동 축소됨
      - Canvas는 픽셀 단위 완전 제어 가능

    주요 기능:
      - 헤더 경계 드래그 → 열 너비 조절
      - 데이터 셀 클릭/드래그 → 범위 선택
      - Ctrl+C → 선택 범위 클립보드 복사
      - Value[hex] 열 더블클릭 → 바이트별 편집 팝업

    상수:
      ROW_H    : 데이터 행 높이 (px)
      HEAD_H   : 헤더 행 높이 (px)
      PAD_X    : 텍스트 좌측 여백 (px)
      SEP_W    : 열 구분선 두께 (px)
      DRAG_TOL : 열 경계 감지 허용 범위 (px)
    """
    ROW_H=24; HEAD_H=28; PAD_X=6; SEP_W=1; DRAG_TOL=5

    def __init__(self,parent,col_cfg,theme,**kw):
        super().__init__(parent,**kw)
        self._col_cfg=col_cfg; self._theme=theme; self.theme=theme
        self._rows=[]; self._col_w=[c[1] for c in col_cfg]
        self._sel_idx=-1; self._dblclick_cb=None; self._scroll_y=0
        self._sel_r0=self._sel_r1=self._sel_c0=self._sel_c1=None
        self._drag_ci=-1; self._drag_x0=0; self._drag_w0=0
        self._col_dragging=False; self._cell_dragging=False
        self._build()

    def _build(self):
        t=self._theme
        self.config(bg=t.get("col_sep","#404040"))
        self._vsb=ttk.Scrollbar(self,orient=tk.VERTICAL,command=self._on_vscroll)
        self._vsb.pack(side=tk.RIGHT,fill=tk.Y)
        self._cv=tk.Canvas(self,highlightthickness=0,bd=0)
        self._cv.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)
        cv=self._cv
        cv.bind("<Configure>",self._on_resize)
        cv.bind("<Motion>",self._on_motion)
        cv.bind("<ButtonPress-1>",self._on_press)
        cv.bind("<B1-Motion>",self._on_drag)
        cv.bind("<ButtonRelease-1>",self._on_release)
        cv.bind("<Double-1>",self._on_dbl)
        cv.bind("<MouseWheel>",self._on_wheel)
        cv.bind("<Button-4>",self._on_wheel)
        cv.bind("<Button-5>",self._on_wheel)
        cv.bind("<Control-c>",self._copy_sel)
        cv.bind("<Control-C>",self._copy_sel)
        cv.config(takefocus=True)
        # 바이트별 편집 팝업 (Value[hex] 더블클릭 시)
        self._byte_popup = None   # tk.Toplevel
        self._byte_entries = []   # 바이트별 Entry 리스트
        self._byte_vals = []      # 현재 바이트 값 리스트
        self._edit_row = -1
        self._edit_cb  = None

    # ── 스크롤 ──────────────────────────────────────────
    def _on_vscroll(self,*args):
        total=len(self._rows)*self.ROW_H+self.HEAD_H
        cv_h=max(self._cv.winfo_height(),1)
        max_y=max(0,total-cv_h)
        if args[0]=="moveto": self._scroll_y=int(float(args[1])*total)
        elif args[0]=="scroll":
            u=self.ROW_H if args[2]=="units" else cv_h
            self._scroll_y+=int(args[1])*u
        self._scroll_y=max(0,min(self._scroll_y,max_y))
        self._upd_sb(); self._redraw()

    def _on_wheel(self,event):
        d=-3 if(event.num==4 or event.delta>0)else 3
        self._scroll_y=max(0,self._scroll_y+d*self.ROW_H)
        self._upd_sb(); self._redraw()

    def _upd_sb(self):
        total=max(1,len(self._rows)*self.ROW_H+self.HEAD_H)
        h=max(self._cv.winfo_height(),1)
        self._vsb.set(self._scroll_y/total,min(1.0,(self._scroll_y+h)/total))

    def _on_resize(self,event): self._recalc_stretch(); self._redraw()

    # ── 열 너비 계산 ─────────────────────────────────────
    # _recalc_stretch: stretch=True 열에 남은 공간 배분
    # fit_col_widths : 데이터 기준 각 열 최적 너비 계산
    def _recalc_stretch(self):
        cw=self._cv.winfo_width()
        if cw<10: return
        fixed=sum(self._col_w[ci] for ci,c in enumerate(self._col_cfg) if not c[4])
        nst=sum(1 for c in self._col_cfg if c[4])
        if nst:
            rem=max(80,cw-fixed-self.SEP_W*len(self._col_cfg))
            sw=rem//nst
            for ci,c in enumerate(self._col_cfg):
                if c[4]: self._col_w[ci]=max(sw,c[2])

    def fit_col_widths(self,rows_data):
        CW=8; PAD=20; n=len(self._col_cfg)
        cw=[len(self._col_cfg[ci][0])*CW+PAD for ci in range(n)]
        for row in rows_data:
            for ci in range(n):
                txt=str(row[ci]) if ci<len(row) else ''
                w=len(txt)*CW+PAD
                if w>cw[ci]: cw[ci]=w
        for ci,c in enumerate(self._col_cfg):
            if not c[4]: self._col_w[ci]=max(cw[ci],c[2])
        self._recalc_stretch()

    # ── 데이터 관리 ──────────────────────────────────────
    # delete_all  : 모든 행 삭제 및 상태 초기화
    # insert_row  : 행 추가 (commit() 전까지 화면 갱신 없음)
    # commit      : insert_row 완료 후 한번 호출 → 화면 갱신
    def delete_all(self):
        self._rows=[]; self._sel_idx=-1; self._scroll_y=0
        self._sel_r0=self._sel_r1=self._sel_c0=self._sel_c1=None
        self._cell_dragging=False
        self._cv.delete("all"); self._upd_sb()

    def insert_row(self,values,tag='even'): self._rows.append((values,tag))

    def commit(self):
        self._recalc_stretch(); self._upd_sb(); self._redraw()

    def get_selected_values(self):
        if 0<=self._sel_idx<len(self._rows): return self._rows[self._sel_idx]
        return None

    def bind_dblclick(self,cb): self._dblclick_cb=cb

    # ── 좌표 헬퍼 ────────────────────────────────────────
    # _col_x  : ci번 열의 시작 x 픽셀 좌표
    # _row_at : y 픽셀 → 행 인덱스 (-1이면 행 없음)
    # _ci_at  : x 픽셀 → 열 인덱스
    # _sep_at : x 픽셀 근처 열 구분선 존재 여부 (열 인덱스 반환)
    def _col_x(self,ci):
        x=0
        for i in range(ci): x+=self._col_w[i]+self.SEP_W
        return x

    def _row_at(self,y):
        ri=(y+self._scroll_y-self.HEAD_H)//self.ROW_H
        return ri if 0<=ri<len(self._rows) else -1

    def _ci_at(self,x):
        for ci in range(len(self._col_cfg)):
            cx=self._col_x(ci)
            if cx<=x<cx+self._col_w[ci]: return ci
        return len(self._col_cfg)-1

    def _sep_at(self,x):
        for ci in range(len(self._col_cfg)-1):
            sx=self._col_x(ci)+self._col_w[ci]
            if abs(x-sx)<=self.DRAG_TOL: return ci
        return -1

    # ── 마우스 이벤트 ─────────────────────────────────────
    # _on_motion  : 커서 모양 변경 (구분선 근처=↔, 데이터=I빔)
    # _on_press   : 헤더=열 크기 드래그 시작 / 데이터=셀 선택 시작
    # _on_drag    : 열 크기 조절 또는 셀 범위 확장
    # _on_release : 드래그 상태 해제
    # _on_dbl     : Value[hex] 열=바이트 팝업 / 나머지=편집바 사용
    def _on_motion(self,event):
        if self._col_dragging: return
        if event.y<=self.HEAD_H and self._sep_at(event.x)>=0:
            self._cv.config(cursor="sb_h_double_arrow")
        elif event.y>self.HEAD_H:
            self._cv.config(cursor="xterm")
        else:
            self._cv.config(cursor="arrow")

    def _on_press(self,event):
        self._cv.focus_set()
        if event.y<=self.HEAD_H:
            ci=self._sep_at(event.x)
            if ci>=0:
                self._drag_ci=ci; self._drag_x0=event.x
                self._drag_w0=self._col_w[ci]; self._col_dragging=True
        else:
            ri=self._row_at(event.y); ci=self._ci_at(event.x)
            if ri>=0:
                self._sel_r0=self._sel_r1=ri
                self._sel_c0=self._sel_c1=ci
                self._sel_idx=ri; self._cell_dragging=True
                self._redraw()

    def _on_drag(self,event):
        if self._col_dragging:
            ci=self._drag_ci; mn=self._col_cfg[ci][2]
            self._col_w[ci]=max(mn,self._drag_w0+event.x-self._drag_x0)
            self._recalc_stretch(); self._redraw()
        elif self._cell_dragging:
            ri=self._row_at(event.y)
            if ri<0: ri=max(0,min(len(self._rows)-1,
                (event.y+self._scroll_y-self.HEAD_H)//self.ROW_H))
            self._sel_r1=ri; self._sel_c1=self._ci_at(event.x)
            self._redraw()

    def _on_release(self,event):
        self._col_dragging=False; self._cell_dragging=False
        self._cv.config(cursor="arrow")

    def _on_dbl(self,event):
        if event.y<self.HEAD_H or self._sep_at(event.x)>=0: return
        ri=self._row_at(event.y)
        if ri<0: return
        self._sel_idx=ri; self._redraw()
        ci=self._ci_at(event.x)
        if ci==2:  # Value[hex] 열 → 인라인 편집
            self._start_inline_edit(ri)
        else:
            if self._dblclick_cb: self._dblclick_cb(ri)

    # ── 바이트 팝업 편집 ──────────────────────────────────
    # _start_inline_edit : Value[hex] 셀 위치에 Toplevel 팝업 생성
    #   - 바이트별 Entry 그리드 (한 줄 8바이트)
    #   - Enter=단일 바이트 확정+다음이동 / Tab=이동 / Esc=취소
    # _byte_commit_one   : 단일 바이트 즉시 반영 (byte_offset 포함)
    # _byte_commit_all   : 전체 바이트 일괄 반영
    def _start_inline_edit(self,ri):
        """바이트별 Entry 팝업 표시"""
        if ri<0 or ri>=len(self._rows): return
        # 기존 팝업 닫기
        self._close_byte_popup()
        vals=self._rows[ri][0]
        cur_val=str(vals[2]) if len(vals)>2 else "00"
        bytes_list=cur_val.split()
        if not bytes_list: return
        self._edit_row=ri
        self._byte_vals=list(bytes_list)
        # 시작 주소 파싱 (Addr(dec) 열: "129~144" 또는 "128")
        try:
            addr_str=str(vals[0]).split("~")[0].strip()
            base_addr=int(addr_str)
        except:
            base_addr=0
        # Canvas 상의 Value열 위치 계산
        col2_x=self._col_x(2)
        ry=self.HEAD_H+ri*self.ROW_H-self._scroll_y
        # 화면 절대 좌표
        rx=self._cv.winfo_rootx()+col2_x
        ry2=self._cv.winfo_rooty()+ry+self.ROW_H
        # Toplevel 팝업 생성
        t=self.theme if hasattr(self,"theme") else {}
        bg=t.get("bg1","#f0f0f0") if t else "#f0f0f0"
        fg_c=t.get("t1","#111111") if t else "#111111"
        acc=t.get("acc","#185FA5") if t else "#185FA5"
        pop=tk.Toplevel(self)
        pop.overrideredirect(True)  # 타이틀바 없음
        pop.config(bg=acc)
        pop.attributes("-topmost",True)
        self._byte_popup=pop
        # 헤더 라벨
        hdr=tk.Frame(pop,bg=acc); hdr.pack(fill=tk.X,padx=1,pady=(1,0))
        end_addr=base_addr+len(bytes_list)-1
        addr_range=f"{base_addr}" if len(bytes_list)==1 else f"{base_addr}~{end_addr}"
        tk.Label(hdr,text=f"Byte 편집  Addr:{addr_range} ({len(bytes_list)}바이트)  Esc=취소",
                 bg=acc,fg="white",font=("",8)).pack(side=tk.LEFT,padx=4)
        tk.Button(hdr,text="✕",bg=acc,fg="white",relief=tk.FLAT,
                  font=("",8),cursor="hand2",
                  command=self._close_byte_popup).pack(side=tk.RIGHT,padx=2)
        # 바이트 Entry 그리드
        gf=tk.Frame(pop,bg=bg); gf.pack(padx=1,pady=(0,1))
        self._byte_entries=[]
        COLS=8  # 한 줄에 8바이트
        for bi,bv in enumerate(bytes_list):
            row_f=tk.Frame(gf,bg=bg)
            if bi%COLS==0: row_f.grid(row=bi//COLS*2+1,column=0,
                                       columnspan=COLS*2,sticky="w")
            # 인덱스 라벨
            col=bi%COLS
            r=bi//COLS
            tk.Label(gf,text=str(base_addr+bi),bg=bg,fg=acc,
                     font=("Consolas",8)).grid(row=r*2,column=col,padx=2,pady=(4,0))
            var=tk.StringVar(value=bv)
            e=tk.Entry(gf,textvariable=var,width=3,
                       font=("Consolas",11),justify="center",
                       bg=bg,fg=fg_c,
                       relief="solid",bd=1,
                       highlightthickness=1,
                       highlightcolor=acc,
                       highlightbackground="#aaaaaa")
            e.grid(row=r*2+1,column=col,padx=2,pady=(0,4))
            e.bind("<Return>",   lambda ev,i=bi: self._byte_commit_one(i))
            e.bind("<KP_Enter>", lambda ev,i=bi: self._byte_commit_one(i))
            e.bind("<Tab>",      lambda ev,i=bi: self._byte_tab(i))
            e.bind("<Escape>",   lambda ev: self._close_byte_popup())
            self._byte_entries.append((e,var))
        # OK 버튼
        bf=tk.Frame(pop,bg=bg); bf.pack(fill=tk.X,padx=4,pady=(0,4))
        tk.Button(bf,text="전체 적용",font=("",9),cursor="hand2",
                  bg=acc,fg="white",relief=tk.FLAT,
                  command=self._byte_commit_all).pack(side=tk.RIGHT,padx=4)
        # 위치 설정
        pop.update_idletasks()
        pw=pop.winfo_reqwidth(); ph=pop.winfo_reqheight()
        # 화면 경계 벗어나지 않도록 조정
        sw=pop.winfo_screenwidth(); sh=pop.winfo_screenheight()
        px=min(rx, sw-pw-4)
        py=min(ry2, sh-ph-4)
        pop.geometry(f"+{px}+{py}")
        # 첫 번째 Entry 포커스
        if self._byte_entries:
            self._byte_entries[0][0].focus_set()
            self._byte_entries[0][0].select_range(0,"end")

    def _close_byte_popup(self,event=None):
        """바이트 팝업 닫기"""
        if self._byte_popup:
            try: self._byte_popup.destroy()
            except: pass
            self._byte_popup=None
        self._byte_entries=[]
        self._edit_row=-1
        self._cv.focus_set()

    def _byte_tab(self,bi):
        """Tab키로 다음 Entry 이동"""
        next_i=bi+1
        if next_i<len(self._byte_entries):
            self._byte_entries[next_i][0].focus_set()
            self._byte_entries[next_i][0].select_range(0,"end")

    def _byte_commit_one(self,bi):
        """단일 바이트 Enter: 현재 바이트만 즉시 반영 후 다음 이동"""
        if bi>=len(self._byte_entries): return
        e,var=self._byte_entries[bi]
        try: val=int(var.get().strip(),16)&0xFF
        except ValueError: return
        # 해당 바이트만 콜백 (offset 포함)
        if self._edit_cb and self._edit_row>=0:
            self._edit_cb(self._edit_row,[val],byte_offset=bi)
        # 다음 Entry로 이동 또는 팝업 닫기
        next_i=bi+1
        if next_i<len(self._byte_entries):
            self._byte_entries[next_i][0].focus_set()
            self._byte_entries[next_i][0].select_range(0,"end")
        else:
            self._close_byte_popup()

    def _byte_commit_all(self):
        """전체 적용 버튼: 모든 바이트 한번에 반영"""
        vals=[]
        for e,var in self._byte_entries:
            try: vals.append(int(var.get().strip(),16)&0xFF)
            except ValueError: vals.append(0)
        if self._edit_cb and self._edit_row>=0:
            self._edit_cb(self._edit_row,vals,byte_offset=0)
        self._close_byte_popup()

    def bind_edit(self,cb):
        """commit 콜백 등록: cb(row_index, new_byte_value)"""
        self._edit_cb=cb

    # ── 클립보드 복사 ─────────────────────────────────────
    # Ctrl+C: 선택된 셀 범위를 탭/개행으로 클립보드 복사
    #         선택 없으면 현재 행 전체 복사
    def _copy_sel(self,event=None):
        r0=self._sel_r0; r1=self._sel_r1
        c0=self._sel_c0; c1=self._sel_c1
        if r0 is None:
            if 0<=self._sel_idx<len(self._rows):
                vals=self._rows[self._sel_idx][0]
                self._clip(chr(9).join(str(v) for v in vals))
            return
        lines=[]
        for ri in range(min(r0,r1),max(r0,r1)+1):
            if ri>=len(self._rows): break
            vals=self._rows[ri][0]
            cells=[str(vals[ci]) if ci<len(vals) else ''
                   for ci in range(min(c0,c1),max(c0,c1)+1)]
            lines.append(chr(9).join(cells))
        self._clip(chr(10).join(lines))

    def _clip(self,text):
        try:
            self.clipboard_clear(); self.clipboard_append(text); self.update()
        except Exception: pass

    # ── 렌더링 ───────────────────────────────────────────
    # _redraw: Canvas 전체를 다시 그림
#   1) 전체 배경 → 2) 헤더(텍스트+구분선) → 3) 데이터 행(배경+텍스트+구분선)
    # 화면 밖 행은 스킵하여 성능 최적화
    def _redraw(self):
        t=self._theme; cv=self._cv
        cv.delete("all")
        cw=cv.winfo_width(); ch=cv.winfo_height()
        if cw<2 or ch<2: return
        n=len(self._col_cfg); sep=t.get('col_sep','#404040')
        cv.create_rectangle(0,0,cw,ch,fill=t.get('bg2','#cccccc'),outline='')
        hbg=t.get('col_head','#b4b4b4'); hfg=t.get('acc','#185FA5')
        cv.create_rectangle(0,0,cw,self.HEAD_H,fill=hbg,outline='')
        for ci in range(n):
            x=self._col_x(ci); w=self._col_w[ci]
            cv.create_text(x+self.PAD_X,self.HEAD_H//2,
                text=self._col_cfg[ci][0],anchor='w',fill=hfg,font=('',10,'bold'))
            if ci<n-1:
                lx=x+w; cv.create_line(lx,0,lx,self.HEAD_H,fill=sep,width=2)
        cv.create_line(0,self.HEAD_H,cw,self.HEAD_H,fill=sep,width=2)
        r0=self._sel_r0; r1=self._sel_r1
        c0=self._sel_c0; c1=self._sel_c1; has_sel=r0 is not None
        a0e=t.get('col0_even','#ccd4e8'); a0o=t.get('col0_odd','#d8deee')
        d1e=t.get('col1_even','#c0c0c0'); d1o=t.get('col1_odd','#cccccc')
        sbg=t.get('sel_bg','#b8d4ee'); sfg=t.get('sel_fg','#0a3a70')
        dfg=t.get('dirty','#7a4000'); fg=t.get('t1','#111111')
        for ri,(vals,tag) in enumerate(self._rows):
            ry=self.HEAD_H+ri*self.ROW_H-self._scroll_y
            if ry+self.ROW_H<0 or ry>ch: continue
            isdirt=(tag=='dirty'); iseven=(ri%2==0)
            row_in_sel=(has_sel and min(r0,r1)<=ri<=max(r0,r1))
            for ci in range(n):
                x=self._col_x(ci); w=self._col_w[ci]
                isaddr=(self._col_cfg[ci][3]==0)
                cell_sel=(row_in_sel and c0 is not None
                          and min(c0,c1)<=ci<=max(c0,c1))
                if cell_sel:    bg=sbg
                elif isdirt:    bg=sbg
                elif isaddr:    bg=a0e if iseven else a0o
                else:           bg=d1e if iseven else d1o
                cv.create_rectangle(x,ry,x+w,ry+self.ROW_H,fill=bg,outline='')
                txt=str(vals[ci]) if ci<len(vals) else ''
                tfg=sfg if cell_sel else(dfg if isdirt else fg)
                tid=cv.create_text(x+self.PAD_X,ry+self.ROW_H//2,
                    text=txt,anchor='w',fill=tfg,font=('Consolas',10))
                bb=cv.bbox(tid)
                if bb and bb[2]>x+w-self.PAD_X:
                    cv.create_rectangle(x+w-1,ry,x+w+self.SEP_W+2,
                        ry+self.ROW_H,fill=bg,outline='')
                if ci<n-1:
                    lx=x+w; cv.create_line(lx,ry,lx,ry+self.ROW_H,fill=sep,width=self.SEP_W)
            cv.create_line(0,ry+self.ROW_H-1,cw,ry+self.ROW_H-1,fill=sep,width=1)

    def apply_theme(self,t):
        self._theme=t; self.config(bg=t.get('col_sep','#404040')); self._redraw()

    def set_col_width(self,ci,w):
        if 0<=ci<len(self._col_w): self._col_w[ci]=w


# ── 메인 앱 ───────────────────────────────────────────────
class App(tk.Tk):
    """
    메인 GUI 애플리케이션.

    상태 관리:
      self.data      : 현재 편집 중인 EEPROM 데이터 {page_key: [0]*128}
      self.orig_data : 파일 로드/EEPROM 읽기 직후 원본 데이터 (dirty 판별용)
      self.conn      : 연결된 I2C 객체 (CP2112I2C 또는 SimulatedI2C)
      self.current_page : 현재 뷰어에 표시 중인 페이지 키

    테마:
      self._dark         : True=다크, False=라이트
      self._bright_dark  : 다크 모드 밝기 (0.7~1.4)
      self._bright_light : 라이트 모드 밝기 (독립 저장)
    """
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME}  v{APP_VERSION}")
        self.geometry("1280x820")
        self.minsize(1000,640)
        self._dark=False
        self._bright_dark=1.0
        self._bright_light=1.0
        self.theme=make_theme(_LIGHT,1.0)
        self.data={k:[0]*128 for k in PAGE_KEYS}
        self.orig_data={k:[0]*128 for k in PAGE_KEYS}
        self.filename=""
        self.conn=None
        self._device_index=0      # 마지막 선택 장치 인덱스
        self._is_simulated=False  # 시뮬레이션 모드 여부
        self.current_page="a0"
        self._build_ui()
        self._apply_theme_full()
        self._refresh_ports()
        # DPI 자동 밝기 조정 비활성화 (기본값 라이트 100% 유지)
        # self.after(300, self._auto_brightness_from_dpi)
        # hidapi 자동 설치 완료 시 팝업 안내
        if HID_OK:
            self.after(800, self._check_hid_installed)

    def _auto_brightness_from_dpi(self):
        try:
            dpi=self.winfo_fpixels("1i")
            f=round(max(0.80,min(1.30,dpi/96.0*0.15+0.86)),2)
            self._bright_dark=f; self._bright_light=f
            self.bright_slider.set(int(f*100))
            self._apply_brightness()
            self._log(f"DPI={dpi:.0f} → 자동 밝기 {int(f*100)}%")
        except Exception as e:
            self._log(f"DPI 감지 실패: {e}")

    # ── UI 빌드 ──────────────────────────────────────────
    def _build_ui(self):
        tb=tk.Frame(self,height=40); tb.pack(fill=tk.X); tb.pack_propagate(False)
        for c,fc in[("●","#ED6A5E"),("●","#F4BF4F"),("●","#61C554")]:
            tk.Label(tb,text=c,fg=fc,font=("",10)).pack(side=tk.LEFT,padx=(8,1),pady=9)
        tk.Label(tb,text=f"  {APP_NAME}",font=("",11,"bold")).pack(side=tk.LEFT,padx=4)
        self.ver_label=tk.Label(tb,text=f"v{APP_VERSION}",font=("",9,"bold"),
                                 fg="#185FA5",cursor="hand2")
        self.ver_label.pack(side=tk.LEFT,padx=2)
        self.ver_label.bind("<Button-1>",lambda e:self._show_ver_dialog())
        tk.Label(tb,text="밝기:").pack(side=tk.RIGHT,padx=(0,3))
        self.bright_slider=tk.Scale(tb,from_=70,to=140,orient=tk.HORIZONTAL,
                                     length=110,showvalue=True,resolution=1,
                                     command=self._on_bright_slide)
        self.bright_slider.set(100)
        self.bright_slider.pack(side=tk.RIGHT,padx=(0,2))
        self.theme_btn=tk.Button(tb,text="다크",width=6,
                                  command=self._toggle_theme,relief=tk.FLAT,cursor="hand2")
        self.theme_btn.pack(side=tk.RIGHT,padx=(0,8))

        self.tab_frame=tk.Frame(self,height=34)
        self.tab_frame.pack(fill=tk.X); self.tab_frame.pack_propagate(False)
        self._tab_btns={}
        for name,label in[("connect","연결"),("viewer","EEPROM 뷰어"),
                           ("write","쓰기"),("log","로그"),("about","정보")]:
            btn=tk.Button(self.tab_frame,text=label,relief=tk.FLAT,bd=0,
                           padx=14,pady=6,cursor="hand2",
                           command=lambda n=name:self._switch_tab(n))
            btn.pack(side=tk.LEFT); self._tab_btns[name]=btn

        self.page_container=tk.Frame(self)
        self.page_container.pack(fill=tk.BOTH,expand=True)
        self._pages={}
        for name in["connect","viewer","write","log","about"]:
            f=tk.Frame(self.page_container); self._pages[name]=f
            {"connect":self._build_connect,"viewer":self._build_viewer,
             "write":self._build_write,"log":self._build_log,
             "about":self._build_about}[name](f)
        self._switch_tab("connect")

    # ════════════════════════════════════════════════════
    # 탭 빌드 메서드
    # _build_connect : CP2112 장치 선택, I2C 속도, 연결 버튼
    # _build_viewer  : 페이지 선택 버튼, CanvasTable, 편집바
    # _build_write   : 쓰기 페이지 선택, 범위, 진행바, 검증
    # _build_log     : 동작 이력 텍스트
    # _build_about   : 버전 히스토리, 시스템 정보
    # ════════════════════════════════════════════════════
    def _build_connect(self,p):
        pad=dict(padx=10,pady=7)
        g1=tk.LabelFrame(p,text=" CP2112-F03-GM USB-HID I2C 브릿지 ",padx=8,pady=8)
        g1.pack(fill=tk.X,padx=16,pady=(16,8))
        tk.Label(g1,text="장치:").grid(row=0,column=0,sticky=tk.W,**pad)
        self.port_var=tk.StringVar()
        self.port_cb=ttk.Combobox(g1,textvariable=self.port_var,width=36,state="readonly")
        self.port_cb.grid(row=0,column=1,sticky=tk.W,**pad)
        tk.Button(g1,text="새로고침",command=self._refresh_ports,width=8,cursor="hand2"
                  ).grid(row=0,column=2,**pad)
        tk.Label(g1,text="I2C 주소:").grid(row=1,column=0,sticky=tk.W,**pad)
        self.i2c_addr_var=tk.StringVar(value="50")  # CMIS 기본값 7bit
        fr=tk.Frame(g1); fr.grid(row=1,column=1,sticky=tk.W,**pad)
        tk.Label(fr,text="0x").pack(side=tk.LEFT)
        tk.Entry(fr,textvariable=self.i2c_addr_var,width=5,
                 font=("Consolas",10)).pack(side=tk.LEFT)
        tk.Label(fr,text="(8-bit hex, 예: A0)").pack(side=tk.LEFT,padx=6)
        tk.Button(g1,text="주소 스캔",width=10,cursor="hand2",
                  command=self._scan_i2c).grid(row=1,column=2,**pad)
        tk.Label(g1,text="I2C 속도:").grid(row=2,column=0,sticky=tk.W,**pad)
        self.i2c_speed_var=tk.StringVar(value="400kHz")
        ttk.Combobox(g1,textvariable=self.i2c_speed_var,width=10,state="readonly",
                     values=["100kHz","400kHz","1MHz"]
                     ).grid(row=2,column=1,sticky=tk.W,**pad)
        self.conn_btn=tk.Button(g1,text="연결",width=10,cursor="hand2",command=self._toggle_connect)
        self.conn_btn.grid(row=0,column=3,rowspan=2,padx=8)
        self.conn_status=tk.Label(g1,text="● 미연결",fg="#f07070",font=("",10,"bold"))
        self.conn_status.grid(row=2,column=3,**pad)
        hid_fg="#2a7a2a" if HID_OK else "#cc3300"
        hid_msg="✓ SLABHIDtoSMBus.dll 로드됨" if HID_OK else "⚠ SLABHIDtoSMBus.dll 없음 (스크립트 폴더에 복사 필요)"
        tk.Label(g1,text=hid_msg,fg=hid_fg,font=("",9)
                 ).grid(row=3,column=0,columnspan=4,sticky=tk.W,padx=10,pady=(0,4))
        g2=tk.LabelFrame(p,text=" 연결 정보 ",padx=8,pady=8)
        g2.pack(fill=tk.X,padx=16,pady=8)
        self.conn_info=tk.Text(g2,height=5,state=tk.DISABLED,relief=tk.FLAT)
        self.conn_info.pack(fill=tk.X)
        g3=tk.LabelFrame(p,text=" CP2112 사용 안내 ",padx=8,pady=8)
        g3.pack(fill=tk.BOTH,expand=True,padx=16,pady=8)
        st=scrolledtext.ScrolledText(g3,height=12,relief=tk.FLAT)
        lines=[
            "[ CP2112-F03-GM USB-HID I2C Bridge ]",
            "",
            "연결 방법:",
            "  1. CP2112 보드를 USB로 PC에 연결",
            "  2. EEPROM SDA/SCL을 CP2112에 연결",
            "  3. I2C 주소 설정 (CMIS Lower Memory: 0x50)",
            "  4. pip install hid",
            "  5. 새로고침 클릭 후 장치 선택 -> 연결",
            "",
            "I2C 주소 (7-bit):",
            "  CMIS Lower Memory (A0h) = 0x50",
            "  CMIS A2h               = 0x51",
            "",
            "드라이버:",
            "  Windows: Silicon Labs CP2112 HID USB-to-SMBus",
            "  Linux  : hid-cp2112 커널 모듈 (내장)",
            "  macOS  : 추가 드라이버 불필요",
            "",
            "설치 확인:",
            "  python -c \"import hid; print(hid.enumerate(0x10C4,0xEA90))\"",
        ]
        st.insert(tk.END,"\n".join(lines))
        st.config(state=tk.DISABLED); st.pack(fill=tk.BOTH,expand=True)


    # ── 뷰어 탭 ──────────────────────────────────────────
    def _build_viewer(self,p):
        top=tk.Frame(p); top.pack(fill=tk.X,padx=12,pady=(10,4))
        for txt,cmd,w in[("파일 열기",self._open_file,9),
                          ("파일 저장",self._save_file,9),
                          ("EEPROM 읽기",self._read_eeprom,11)]:
            tk.Button(top,text=txt,width=w,cursor="hand2",command=cmd
                      ).pack(side=tk.LEFT,padx=(0,5))
        self.file_label=tk.Label(top,text="파일 없음",anchor=tk.W)
        self.file_label.pack(side=tk.LEFT,padx=8)
        self.dirty_label=tk.Label(top,text="")
        self.dirty_label.pack(side=tk.LEFT)

        pg=tk.Frame(p); pg.pack(fill=tk.X,padx=12,pady=(0,4))
        tk.Label(pg,text="페이지:").pack(side=tk.LEFT,padx=(0,6))
        self._page_btns={}
        for pk,lbl in zip(PAGE_KEYS,PAGE_LABELS):
            btn=tk.Button(pg,text=lbl,width=8,cursor="hand2",
                           command=lambda k=pk:self._show_page(k))
            btn.pack(side=tk.LEFT,padx=2); self._page_btns[pk]=btn

        # CanvasTable
        self.mct=CanvasTable(p, COL_CFG, self.theme)
        self.mct.pack(fill=tk.BOTH,expand=True,padx=12,pady=(0,4))
        self.mct.bind_dblclick(self._on_row_dblclick)
        self.mct.bind_edit(self._on_inline_edit)

        edit=tk.Frame(p); edit.pack(fill=tk.X,padx=12,pady=(0,6))
        tk.Label(edit,text="편집 (hex):").pack(side=tk.LEFT)
        self.edit_var=tk.StringVar()
        self.edit_entry=tk.Entry(edit,textvariable=self.edit_var,width=8)
        self.edit_entry.pack(side=tk.LEFT,padx=4)
        tk.Button(edit,text="적용",width=6,cursor="hand2",command=self._apply_edit
                  ).pack(side=tk.LEFT)
        self.edit_info=tk.Label(edit,text="Value[hex] 더블클릭: 직접 편집 / 다른 열 더블클릭: 하단 편집",anchor=tk.W)
        self.edit_info.pack(side=tk.LEFT,padx=8)

    # ── 쓰기 탭 ──────────────────────────────────────────
    def _build_write(self,p):
        g=tk.LabelFrame(p,text=" 쓰기 페이지 선택 ",padx=8,pady=8)
        g.pack(fill=tk.X,padx=16,pady=(16,8))
        self.write_vars={pk:tk.BooleanVar(value=(pk in["a0","p00"])) for pk in PAGE_KEYS}
        for i,(pk,lbl) in enumerate(zip(PAGE_KEYS,PAGE_LABELS)):
            tk.Checkbutton(g,text=lbl,variable=self.write_vars[pk]).grid(
                row=i//3,column=i%3,sticky=tk.W,padx=12,pady=4)
        g2=tk.LabelFrame(p,text=" 쓰기 범위 ",padx=8,pady=8)
        g2.pack(fill=tk.X,padx=16,pady=8)
        tk.Label(g2,text="시작(hex):").grid(row=0,column=0,sticky=tk.W,padx=8,pady=4)
        self.write_start=tk.Entry(g2,width=8); self.write_start.insert(0,"00")
        self.write_start.grid(row=0,column=1,sticky=tk.W,padx=4)
        tk.Label(g2,text="끝(hex):").grid(row=0,column=2,sticky=tk.W,padx=8)
        self.write_end=tk.Entry(g2,width=8); self.write_end.insert(0,"7F")
        self.write_end.grid(row=0,column=3,sticky=tk.W,padx=4)
        tk.Label(g2,text="수정 바이트만:").grid(row=0,column=4,sticky=tk.W,padx=8)
        self.write_dirty_only=tk.BooleanVar(value=False)
        tk.Checkbutton(g2,variable=self.write_dirty_only).grid(row=0,column=5)
        g3=tk.LabelFrame(p,text=" 진행 상태 ",padx=8,pady=8)
        g3.pack(fill=tk.X,padx=16,pady=8)
        self.progress_var=tk.DoubleVar()
        ttk.Progressbar(g3,variable=self.progress_var,maximum=100).pack(fill=tk.X,pady=4)
        self.progress_label=tk.Label(g3,text="대기 중"); self.progress_label.pack()
        bf=tk.Frame(p); bf.pack(pady=12)
        self.write_btn=tk.Button(bf,text="EEPROM 쓰기 시작",width=18,
                                  font=("",11,"bold"),cursor="hand2",command=self._start_write)
        self.write_btn.pack(side=tk.LEFT,padx=8)
        tk.Button(bf,text="쓰기 검증 (Read-back)",width=18,cursor="hand2",
                  command=self._verify_write).pack(side=tk.LEFT,padx=8)
        g4=tk.LabelFrame(p,text=" 쓰기 예정 요약 ",padx=8,pady=8)
        g4.pack(fill=tk.BOTH,expand=True,padx=16,pady=(0,4))
        self.write_summary=scrolledtext.ScrolledText(g4,height=8,relief=tk.FLAT)
        self.write_summary.pack(fill=tk.BOTH,expand=True)
        tk.Button(p,text="요약 새로고침",cursor="hand2",command=self._refresh_summary
                  ).pack(pady=(0,8))

    # ── 로그 탭 ──────────────────────────────────────────
    def _build_log(self,p):
        bar=tk.Frame(p); bar.pack(fill=tk.X,padx=12,pady=8)
        tk.Button(bar,text="로그 지우기",cursor="hand2",command=self._clear_log
                  ).pack(side=tk.LEFT)
        tk.Button(bar,text="로그 저장",cursor="hand2",command=self._save_log
                  ).pack(side=tk.LEFT,padx=6)
        self.log_text=scrolledtext.ScrolledText(p,relief=tk.FLAT,wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH,expand=True,padx=12,pady=(0,12))
        self.log_text.config(state=tk.DISABLED)

    # ── 정보 탭 ──────────────────────────────────────────
    def _build_about(self,p):
        hf=tk.Frame(p); hf.pack(fill=tk.X,padx=24,pady=(24,8))
        tk.Label(hf,text=APP_NAME,font=("",16,"bold")).pack(anchor=tk.W)
        tk.Label(hf,text=f"Version {APP_VERSION}  ({APP_DATE})",font=("",10)
                 ).pack(anchor=tk.W,pady=2)
        tk.Label(hf,text="CMIS 5.0 기반 EEPROM 관리 도구  |  CP2112 USB-HID 지원",
                 font=("",9)).pack(anchor=tk.W,pady=4)
        lf=tk.Frame(p); lf.pack(anchor=tk.W,padx=24,pady=(0,4))
        for typ,badge in TYPE_BADGE.items():
            tk.Label(lf,text=f"{badge} = {typ}",font=("",9),
                     fg=TYPE_COLOR[typ]).pack(side=tk.LEFT,padx=(0,16))
        ttk.Separator(p,orient=tk.HORIZONTAL).pack(fill=tk.X,padx=24,pady=6)
        tk.Label(p,text="변경 이력",font=("",11,"bold")).pack(anchor=tk.W,padx=24,pady=(0,4))
        vf=tk.Frame(p); vf.pack(fill=tk.BOTH,expand=True,padx=24,pady=(0,8))
        vt=ttk.Treeview(vf,columns=("typ","ver","date","note"),show="headings",
                          height=8,selectmode="none")
        for c,h,w in zip(("typ","ver","date","note"),["","버전","날짜","변경 내용"],
                          [36,80,100,520]):
            vt.heading(c,text=h); vt.column(c,width=w,minwidth=20,stretch=(c=="note"))
        for ver,date,typ,note in CHANGELOG:
            tag="latest" if ver==APP_VERSION else typ
            vt.insert("",tk.END,values=(TYPE_BADGE.get(typ,""),ver,date,note),tags=(tag,))
        vsb2=ttk.Scrollbar(vf,orient=tk.VERTICAL,command=vt.yview)
        vt.configure(yscrollcommand=vsb2.set)
        vt.pack(side=tk.LEFT,fill=tk.BOTH,expand=True); vsb2.pack(side=tk.RIGHT,fill=tk.Y)
        self._about_tree=vt
        ttk.Separator(p,orient=tk.HORIZONTAL).pack(fill=tk.X,padx=24,pady=4)
        sf=tk.Frame(p); sf.pack(fill=tk.X,padx=24,pady=(4,16))
        pys="로드됨" if HID_OK else "없음 (스크립트 폴더에 DLL 복사 필요)"
        for lbl,val in[("Python",sys.version.split()[0]),
                        ("SLABHIDtoSMBus.dll",pys),("플랫폼",sys.platform)]:
            r=tk.Frame(sf); r.pack(fill=tk.X,pady=1)
            tk.Label(r,text=lbl+":",width=10,anchor=tk.W,font=("",9)).pack(side=tk.LEFT)
            tk.Label(r,text=val,anchor=tk.W,font=("",9)).pack(side=tk.LEFT)

    # ── 버전 팝업 ────────────────────────────────────────
    def _show_ver_dialog(self):
        d=tk.Toplevel(self); d.title("버전 정보")
        d.geometry("540x340"); d.resizable(False,False); d.grab_set()
        t=self.theme; d.config(bg=t["bg1"])
        tk.Label(d,text=APP_NAME,font=("",14,"bold"),bg=t["bg1"],fg=t["t1"]
                 ).pack(pady=(18,3))
        tk.Label(d,text=f"Version {APP_VERSION}  ·  {APP_DATE}",font=("",10),
                 bg=t["bg1"],fg=t["t2"]).pack()
        lf2=tk.Frame(d,bg=t["bg1"]); lf2.pack(pady=(4,2))
        for typ,badge in TYPE_BADGE.items():
            tk.Label(lf2,text=f"{badge}={typ}",font=("",9),
                     bg=t["bg1"],fg=TYPE_COLOR[typ]).pack(side=tk.LEFT,padx=8)
        ttk.Separator(d,orient=tk.HORIZONTAL).pack(fill=tk.X,padx=20,pady=8)
        fr=tk.Frame(d,bg=t["bg1"]); fr.pack(fill=tk.BOTH,expand=True,padx=20,pady=(0,8))
        vt=ttk.Treeview(fr,columns=("typ","ver","date","note"),show="headings",
                          height=7,selectmode="none")
        for c,h,w in zip(("typ","ver","date","note"),["","버전","날짜","내용"],
                          [36,72,92,280]):
            vt.heading(c,text=h); vt.column(c,width=w,stretch=(c=="note"))
        for ver,date,typ,note in CHANGELOG:
            tag="latest" if ver==APP_VERSION else typ
            vt.insert("",tk.END,values=(TYPE_BADGE.get(typ,""),ver,date,note),tags=(tag,))
        vt.tag_configure("latest",foreground=t["acc"],font=("Consolas",10,"bold"))
        vt.tag_configure("major",foreground=t["red"])
        vt.tag_configure("minor",foreground=t["acc"])
        vt.tag_configure("patch",foreground=t["t3"])
        vsb3=ttk.Scrollbar(fr,orient=tk.VERTICAL,command=vt.yview)
        vt.configure(yscrollcommand=vsb3.set)
        vt.pack(side=tk.LEFT,fill=tk.BOTH,expand=True); vsb3.pack(side=tk.RIGHT,fill=tk.Y)
        tk.Button(d,text="닫기",width=10,command=d.destroy,cursor="hand2",
                  bg=t["btn_bg"],fg=t["btn_fg"],relief=tk.FLAT).pack(pady=(0,14))

    # ── 탭 전환 ──────────────────────────────────────────
    def _switch_tab(self,name):
        for n,f in self._pages.items(): f.place_forget()
        self._pages[name].place(relx=0,rely=0,relwidth=1,relheight=1)
        t=self.theme
        for n,btn in self._tab_btns.items():
            active=(n==name)
            btn.config(fg=t["acc"] if active else t["t2"],
                       bg=t["bg2"] if active else t["bg1"],relief=tk.FLAT)

    # ── 밝기 / 테마 ──────────────────────────────────────
    def _on_bright_slide(self,val):
        f=int(val)/100.0
        if self._dark: self._bright_dark=f
        else:          self._bright_light=f
        self._apply_brightness()

    def _apply_brightness(self):
        f=self._bright_dark if self._dark else self._bright_light
        base=_DARK if self._dark else _LIGHT
        self.theme=make_theme(base,f)
        self._apply_theme_full()

    def _toggle_theme(self):
        self._dark=not self._dark
        self.theme_btn.config(text="라이트" if self._dark else "다크")
        f=self._bright_dark if self._dark else self._bright_light
        self.bright_slider.set(int(f*100))
        self._apply_brightness()

    def _apply_theme_full(self):
        t=self.theme
        self.config(bg=t["bg0"])
        self._walk(self,t)
        self._style_misc(t)
        try: self.bright_slider.config(bg=t["bg1"],fg=t["t1"],
                                        troughcolor=t["bg2"],
                                        highlightbackground=t["bg1"])
        except: pass
        try: self.mct.apply_theme(t)
        except: pass
        cur=[n for n,f in self._pages.items() if str(f.place_info())!="{}"]
        if cur: self._switch_tab(cur[0])

    def _walk(self,widget,t):
        cls=widget.winfo_class()
        try:
            if cls in("Frame","LabelFrame","Toplevel"):
                widget.config(bg=t["bg1"])
            elif cls=="Label":
                # ver_label은 acc 색상 고정 (버전 표시)
                if widget is getattr(self,"ver_label",None):
                    widget.config(bg=t["bg1"],fg=t["acc"])
                else:
                    widget.config(bg=t["bg1"],fg=t["t1"])
            elif cls=="Button":
                widget.config(bg=t["btn_bg"],fg=t["btn_fg"],
                               activebackground=t["bg0"],relief=tk.FLAT)
            elif cls=="Entry":
                widget.config(bg=t["entry_bg"],fg=t["entry_fg"],
                               insertbackground=t["t1"],relief=tk.FLAT,
                               highlightthickness=1,highlightcolor=t["acc"],
                               highlightbackground=t["bd"])
            elif cls=="Text":
                widget.config(bg=t["bg2"],fg=t["t1"],
                               insertbackground=t["t1"],relief=tk.FLAT)
            elif cls=="Checkbutton":
                widget.config(bg=t["bg1"],fg=t["t1"],
                               activebackground=t["bg1"],selectcolor=t["bg3"])
        except: pass
        for child in widget.winfo_children():
            self._walk(child,t)

    def _style_misc(self,t):
        s=ttk.Style(self); s.theme_use("default")
        s.configure("TCombobox",fieldbackground=t["entry_bg"],background=t["bg2"],
                    foreground=t["t1"],selectbackground=t["sel_bg"])
        s.configure("Vertical.TScrollbar",background=t["bg1"],
                    troughcolor=t["bg0"],arrowcolor=t["t3"])
        s.configure("Horizontal.TScrollbar",background=t["bg1"],
                    troughcolor=t["bg0"],arrowcolor=t["t3"])
        s.configure("TProgressbar",background=t["acc"],troughcolor=t["bg2"])
        # 정보탭 버전 트리
        s.configure("Treeview",background=t["bg2"],foreground=t["t1"],
                    fieldbackground=t["bg2"],rowheight=22,font=("Consolas",10))
        s.configure("Treeview.Heading",background=t["col_head"],foreground=t["acc"],
                    relief="groove",borderwidth=2,font=("",10,"bold"))
        s.map("Treeview",background=[("selected",t["sel_bg"])],
              foreground=[("selected",t["sel_fg"])])
        try:
            self._about_tree.tag_configure("latest",foreground=t["acc"],
                                            font=("Consolas",10,"bold"))
            self._about_tree.tag_configure("major",foreground=t["red"])
            self._about_tree.tag_configure("minor",foreground=t["acc"])
            self._about_tree.tag_configure("patch",foreground=t["t3"])
        except: pass

    # ── 포트 ─────────────────────────────────────────────
    def _scan_i2c(self):
        """I2C 주소 스캔: 포트 열어서 탐색 후 닫기"""
        if self._device_index < 0 and not self._is_simulated:
            messagebox.showwarning("스캔 불가","먼저 연결하세요.")
            return
        self._log("I2C 주소 스캔 시작 (0x08~0x77)...")
        self.conn_status.config(text="● 스캔 중...",fg=self.theme["amb"])
        self.update()
        def task():
            conn=None
            try:
                conn=self._open_conn()
                found=conn.scan_addresses()
                self.after(0,lambda f=found: self._scan_done(f))
            except Exception as e:
                logging.error("I2C 스캔 오류",exc_info=True)
                self.after(0,lambda e=e: messagebox.showerror("스캔 오류",str(e)))
            finally:
                self._close_conn(conn)
        threading.Thread(target=task,daemon=True).start()

    def _scan_done(self,found):
        """스캔 완료 처리: 결과 표시 및 주소 자동 설정"""
        self.conn_status.config(text="● CP2112 연결됨",fg=self.theme["grn"])
        if not found:
            self._log("스캔 결과: 응답 장치 없음")
            messagebox.showinfo("스캔 결과","응답하는 I2C 장치가 없습니다.\n배선/전원/풀업저항을 확인하세요.")
            return
        addr_list="\n".join(f"  0x{a:02X} ({a})" for a in found)
        self._log(f"스캔 결과: {[hex(a) for a in found]}")
        # 첫 번째 발견 주소로 자동 설정
        self.i2c_addr_var.set(f"{found[0]:02X}")
        msg=(f"발견된 I2C 주소:\n{addr_list}\n\n"
             f"0x{found[0]:02X}으로 자동 설정했습니다.")
        messagebox.showinfo("스캔 완료",msg)
        self._update_conn_info()

    def _check_hid_installed(self):
        """자동 설치 후 안내 팝업 (최초 1회)"""
        # 연결 탭의 hid 상태 레이블 갱신
        try:
            self._refresh_ports()
        except Exception:
            pass

    def _refresh_ports(self):
        if not HID_OK:
            self.port_cb["values"]=["SLABHIDtoSMBus.dll 없음"]; return
        devs=CP2112I2C.list_devices()
        if devs:
            items=[f"[{i}] {d.get('product_string','CP2112')} S/N:{d.get('serial_number','')}"
                   for i,d in enumerate(devs)]
            self.port_cb["values"]=items; self.port_var.set(items[0])
        else:
            self.port_cb["values"]=["CP2112 장치 없음 (USB 연결 확인)"]


    def _toggle_connect(self):
        # 연결 해제
        if self._device_index >= 0 and self.conn_btn.cget("text") == "연결 해제":
            if self.conn:
                self.conn.close(); self.conn=None
            self._device_index = -1
            self._is_simulated = False
            self.conn_status.config(text="● 미연결",fg=self.theme["red"])
            self.conn_btn.config(text="연결")
            self._log("연결 해제"); return
        # 시뮬레이션 모드
        if not HID_OK:
            if messagebox.askyesno("DLL 없음",
                "SLABHIDtoSMBus.dll을 스크립트 폴더에 복사하세요.\n\n시뮬레이션 모드로 연결?"):
                self._is_simulated = True
                self._device_index = -1
                self.conn_status.config(text="● 시뮬레이션",fg=self.theme["amb"])
                self.conn_btn.config(text="연결 해제")
                self._log("시뮬레이션 모드"); self._update_conn_info()
            return
        # 장치 인덱스 파싱 후 저장 (포트는 실제 사용 시에만 열기)
        sel=self.port_var.get()
        try: idx=int(sel.split("]")[0].replace("[","").strip())
        except: idx=0
        # 연결 테스트 (정상이면 즉시 닫고 인덱스만 기억)
        try:
            test=CP2112I2C(device_index=idx)
            test.close()
            self._device_index=idx
            self._is_simulated=False
            self.conn=None  # 포트는 작업 시에만 점유
            self.conn_status.config(text="● 대기 중 (포트 해제됨)",fg=self.theme["grn"])
            self.conn_btn.config(text="연결 해제")
            self._log(f"CP2112 장치[{idx}] 확인됨 - 읽기/쓰기 시 포트 자동 점유")
            self._update_conn_info()
        except Exception as e:
            logging.error("CP2112 연결 실패", exc_info=True)
            messagebox.showerror("연결 실패",str(e)); self._log(f"연결 실패: {e}")


    def _open_conn(self):
        """작업 전 포트 열기. 시뮬레이션이면 SimulatedI2C 반환."""
        if self._is_simulated:
            return SimulatedI2C()
        if self._device_index < 0:
            raise IOError("장치가 선택되지 않았습니다. 먼저 연결하세요.")
        return CP2112I2C(device_index=self._device_index)

    def _close_conn(self, conn):
        """작업 후 포트 닫기."""
        try:
            if conn: conn.close()
        except: pass
        self.conn=None
        if not self._is_simulated:
            self.after(0, lambda: self.conn_status.config(
                text="● 대기 중 (포트 해제됨)", fg=self.theme["grn"]))

    def _update_conn_info(self):
        self.conn_info.config(state=tk.NORMAL); self.conn_info.delete("1.0",tk.END)
        if self._is_simulated:
            txt=("장치: 시뮬레이션\n"
                 "I2C 주소: 0x"+self.i2c_addr_var.get()+" (7-bit)\n"
                 "상태: 시뮬레이션\n"
                 "※ SLABHIDtoSMBus.dll을 스크립트 폴더에 복사 후 재시작")
        else:
            devs=CP2112I2C.list_devices()
            sel=self.port_var.get()
            try: idx=int(sel.split("]")[0].replace("[","").strip())
            except: idx=0
            d=devs[idx] if idx<len(devs) else {}
            txt=(f"장치: {d.get('product_string','CP2112')}\n"
                 f"S/N : {d.get('serial_number','N/A')}\n"
                 f"I2C 주소: 0x{self.i2c_addr_var.get()} (7-bit)\n"
                 f"I2C 속도: {self.i2c_speed_var.get()}\n"
                 "상태: 연결됨")
        self.conn_info.insert(tk.END,txt)
        self.conn_info.config(state=tk.DISABLED)


    # ── 파일 I/O ─────────────────────────────────────────
    def _open_file(self):
        path=filedialog.askopenfilename(filetypes=[("Text","*.txt"),("All","*.*")])
        if not path: return
        with open(path,"r",encoding="utf-8") as f: self.data=parse_txt(f.read())
        self.orig_data=copy.deepcopy(self.data)
        self.filename=path; self.file_label.config(text=os.path.basename(path))
        self._show_page(self.current_page); self._log(f"파일 로드: {path}")

    def _save_file(self):
        base=(os.path.splitext(os.path.basename(self.filename))[0]
              if self.filename else "eeprom")
        path=filedialog.asksaveasfilename(
            defaultextension=".txt",initialfile=base+"_edited.txt",
            filetypes=[("Text","*.txt"),("All","*.*")])
        if not path: return
        with open(path,"w",encoding="utf-8") as f: f.write(build_txt(self.data))
        self._log(f"파일 저장: {path}")

    # ── 뷰어 ─────────────────────────────────────────────
    # ── 뷰어 데이터 갱신 ─────────────────────────────────
    # _show_page   : 페이지 전환 시 CanvasTable 전체 재구성
    #   1) 페이지 버튼 강조 갱신
    #   2) mct.delete_all() → _get_rows() → mct.insert_row()
    #   3) fit_col_widths() → commit()
    # _get_rows    : 페이지 데이터 → 6컬럼 tuple 리스트
    # _decoded_a0  : A0h 바이트 → Decoded 열 문자열
    # _desc_a0     : A0h 바이트 → Description 열 한국어 설명
    # _decoded_p00 : P00h 바이트 → Decoded 열 문자열
    # _desc_p00    : P00h 바이트 → Description 열 한국어 설명
    # _is_dirty    : 원본 대비 변경 여부 (dirty=파란 배경+주황 텍스트)
    def _show_page(self,pk):
        self.current_page=pk
        t=self.theme
        for k,btn in self._page_btns.items():
            btn.config(fg=t["acc"] if k==pk else t["t2"],
                       bg=t["sel_bg"] if k==pk else t["btn_bg"])
        self.mct.delete_all()
        rows=self._get_rows(pk)
        for i,row in enumerate(rows):
            dirty=self._is_dirty(pk,row[1])
            tag="dirty" if dirty else ("even" if i%2==0 else "odd")
            self.mct.insert_row(row, tag)
        self.mct.fit_col_widths(rows)
        self.mct.commit()
        self._update_dirty_label()

    def _get_rows(self,pk):
        v=self.data[pk]; rows=[]
        if pk=="a0":
            defs=[
                (0,1,"SFF8024Identifier"),(1,1,"CmisRevision"),(2,1,"MemoryModel"),
                (3,1,"GlobalStatus"),
                (4,1,"FlagsSummary Bank0"),(5,1,"FlagsSummary Bank1"),
                (6,1,"FlagsSummary Bank2"),(7,1,"FlagsSummary Bank3"),
                (8,1,"ModuleFlags[CDB/FW]"),(9,1,"ModuleFlags[Vcc/Temp]"),
                (10,1,"ModuleFlags[0A]"),(11,1,"ModuleFlags[0B]"),
                (12,1,"ModuleFlags[0C]"),(13,1,"ModuleFlags[0D]"),
                (14,2,"TempMonValue"),(16,2,"VccMonVoltage"),
                (18,2,"Aux1MonValue"),(20,2,"Aux2MonValue"),
                (22,2,"Aux3MonValue"),(24,2,"CustomMonValue"),
                (26,1,"ModuleControls"),(37,1,"CdbStatus1"),
                (38,1,"CdbStatus2"),(39,2,"ActiveFWVersion"),
                (85,1,"AppDesc[0] HostIfID"),(86,1,"AppDesc[0] MediaIfID"),
                (87,1,"AppDesc[0] LaneCnt"),(88,1,"AppDesc[0] LaneAssign"),
                (126,1,"PageSelectBank"),(127,1,"PageSelectPage"),
            ]
        elif pk=="p00":
            defs=[
                (0,1,"SFF8024IdentifierCopy"),(1,16,"VendorName"),
                (17,3,"VendorOUI"),(20,16,"VendorPN"),(36,2,"VendorRev"),
                (38,16,"VendorSN"),(54,8,"DateCode"),(62,10,"CLEICode"),
                (72,1,"ModulePowerClass"),(73,1,"MaxPower"),
                (74,1,"CableLinkLength"),(75,1,"ConnectorType"),
                (84,1,"MediaInterfaceTech"),(94,1,"PageChecksum"),
            ]
        elif pk=="p01":
            defs=[
                (0,2,"InactiveFWVersion"),(2,2,"HardwareVersion"),
                (4,2,"LengthSMF"),(6,1,"LengthOM5"),(7,1,"LengthOM4"),
                (8,1,"LengthOM3"),(9,1,"LengthOM2"),
                (10,2,"NominalWavelength"),(12,2,"WavelengthTolerance"),
                (14,1,"SupportedPages"),(15,2,"DurationAdvertising"),
                (17,10,"ModuleCharacteristics"),(27,2,"SupportedControls"),
                (29,2,"SupportedFlags"),(31,2,"SupportedMonitors"),
                (33,2,"SupportedSIControls"),(35,4,"SupportedCDB"),
                (127,1,"PageChecksum"),
            ]
        elif pk=="p02":
            defs=[
                (0,2,"TempHighAlarm"),(2,2,"TempLowAlarm"),
                (4,2,"TempHighWarning"),(6,2,"TempLowWarning"),
                (8,2,"VccHighAlarm"),(10,2,"VccLowAlarm"),
                (12,2,"VccHighWarning"),(14,2,"VccLowWarning"),
                (16,2,"Aux1HighAlarm"),(18,2,"Aux1LowAlarm"),
                (20,2,"Aux1HighWarning"),(22,2,"Aux1LowWarning"),
                (24,2,"Aux2HighAlarm"),(26,2,"Aux2LowAlarm"),
                (28,2,"Aux2HighWarning"),(30,2,"Aux2LowWarning"),
                (32,2,"Aux3HighAlarm"),(34,2,"Aux3LowAlarm"),
                (36,2,"Aux3HighWarning"),(38,2,"Aux3LowWarning"),
                (40,2,"CustomHighAlarm"),(42,2,"CustomLowAlarm"),
                (44,2,"CustomHighWarning"),(46,2,"CustomLowWarning"),
                (48,2,"TxPwrHighAlarm"),(50,2,"TxPwrLowAlarm"),
                (52,2,"TxPwrHighWarning"),(54,2,"TxPwrLowWarning"),
                (56,2,"TxBiasHighAlarm"),(58,2,"TxBiasLowAlarm"),
                (60,2,"TxBiasHighWarning"),(62,2,"TxBiasLowWarning"),
                (64,2,"RxPwrHighAlarm"),(66,2,"RxPwrLowAlarm"),
                (68,2,"RxPwrHighWarning"),(70,2,"RxPwrLowWarning"),
                (127,1,"PageChecksum"),
            ]
        else:
            base=0x80
            for i in range(128):
                if v[i]:
                    rows.append((str(i+base),f"{i+base:02X}h",h2(v[i]),
                                  h2(v[i]),f"{pk.upper()}:{i+base:02X}h",""))
            return rows
        base=0x80 if pk!="a0" else 0
        for start,ln,field in defs:
            ad=(str(start+base) if ln==1 else f"{start+base}~{start+ln-1+base}")
            ah=(f"{start+base:02X}h" if ln==1 else
                f"{start+base:02X}~{start+ln-1+base:02X}h")
            # Value[hex]: 전체 바이트 표시
            vs=" ".join(h2(v[start+i]) for i in range(ln))
            if pk=="a0":
                decoded=self._decoded_a0(v,start)
                desc=self._desc_a0(v,start)
            elif pk=="p00":
                decoded=self._decoded_p00(v,start)
                desc=self._desc_p00(v,start)
            elif pk=="p01":
                decoded=self._decoded_p01(v,start)
                desc=self._desc_p01(start)
            elif pk=="p02":
                decoded=self._decoded_p02(v,start)
                desc=self._desc_p02(start)
            else:
                decoded=""; desc=""
            rows.append((ad,ah,vs,decoded,field,desc))
        return rows

    def _decoded_a0(self,v,i):
        """Value 열 옆 Decoded 열: 숫자/상태를 간결하게"""
        if i==0:  return SFF8024.get(v[0],f"?({h2(v[0])})")
        if i==1:  return f"v{(v[1]>>4)&0xF}.{v[1]&0xF}"
        if i==2:
            mci=["≤400kHz","≤1MHz","Rsvd","Rsvd"][(v[2]>>2)&3]
            return f"{'Flat' if v[2]>>7 else 'Paged'} MCI={mci}"
        if i==3:
            ms=(v[3]>>1)&7
            st=["Rsvd","LowPwr","PwrUp","Ready","PwrDn","Fault","Rsvd","Rsvd"]
            return f"{st[ms]} Int={'No' if v[3]&1 else 'ASSERT'}"
        if i>=4 and i<=7:
            b=i-4; flags=[]
            if v[i]&8: flags.append(f"B{b}P2C")
            if v[i]&4: flags.append(f"B{b}P14")
            if v[i]&2: flags.append(f"B{b}P12")
            if v[i]&1: flags.append(f"B{b}P11")
            return ",".join(flags) if flags else "—"
        if i==8:
            fl=[]
            if v[8]&0x80: fl.append("CdbCmp2")
            if v[8]&0x40: fl.append("CdbCmp1")
            if v[8]&0x04: fl.append("DPFWErr")
            if v[8]&0x02: fl.append("ModFWErr")
            if v[8]&0x01: fl.append("StateChg")
            return ",".join(fl) if fl else "—"
        if i==9:
            fl=[]
            if v[9]&0x80: fl.append("VccLowW")
            if v[9]&0x40: fl.append("VccHiW")
            if v[9]&0x20: fl.append("VccLowA")
            if v[9]&0x10: fl.append("VccHiA")
            if v[9]&0x08: fl.append("TmpLowW")
            if v[9]&0x04: fl.append("TmpHiW")
            if v[9]&0x02: fl.append("TmpLowA")
            if v[9]&0x01: fl.append("TmpHiA")
            return ",".join(fl) if fl else "N/A"
        if i==14:
            raw=v[14]<<8|v[15]
            t=(raw-65536 if raw>=0x8000 else raw)/256
            return f"{t:.2f} °C"
        if i==16:
            return f"{(v[16]<<8|v[17])*100/1e6:.4f} V"
        if i==18: return f"S16={v[18]<<8|v[19]}"
        if i==20: return f"S16={v[20]<<8|v[21]}"
        if i==22: return f"S16={v[22]<<8|v[23]}"
        if i==24: return f"S16={v[24]<<8|v[25]}"
        if i==26:
            return (f"LowPwrHW={(v[26]>>6)&1} "
                    f"Squelch={(v[26]>>5)&1} "
                    f"LowPwrSW={(v[26]>>4)&1} "
                    f"Reset={(v[26]>>3)&1}")
        if i==37: return f"Busy={(v[37]>>7)&1} Fail={(v[37]>>6)&1} Res={h2(v[37]&0x3F)}"
        if i==38: return f"Busy={(v[38]>>7)&1} Fail={(v[38]>>6)&1} Res={h2(v[38]&0x3F)}"
        if i==39: return f"v{v[39]}.{v[40]}"
        if i==85: return f"HostIfID={h2(v[85])}"
        if i==86: return f"MediaIfID={h2(v[86])}"
        if i==87: return f"Host={v[87]>>4} Media={v[87]&0xF}"
        if i==88: return f"LaneAssign={h2(v[88])}"
        if i==126: return f"Bank={v[126]}"
        if i==127: return f"Page={h2(v[127])}"
        return "—"

    def _desc_a0(self,v,i):
        """Description 열: 필드 의미 설명"""
        if i==0:  return "SFF-8024 모듈 타입 식별자"
        if i==1:  return "CMIS 규격 버전 (상위4=Major, 하위4=Minor)"
        if i==2:  return "메모리 모델 / MCI 최대 속도"
        if i==3:  return "모듈 현재 상태 / 인터럽트 신호"
        if i>=4 and i<=7: return f"Bank{i-4} 페이지별 Flag 요약"
        if i==8:  return "CDB 완료 / 펌웨어 오류 Flags"
        if i==9:  return "Vcc/온도 알람·경고 Flags"
        if i==14: return "모듈 온도 (S16, 1/256°C 단위)"
        if i==16: return "공급 전압 (U16, 100µV 단위)"
        if i==18: return "Aux1 모니터 (TEC 전류 또는 Custom)"
        if i==20: return "Aux2 모니터 (레이저 온도 또는 TEC)"
        if i==22: return "Aux3 모니터 (레이저 온도 또는 추가 전압)"
        if i==24: return "Custom 모니터"
        if i==26: return "모듈 전역 제어 비트"
        if i==37: return "CDB 인스턴스1 명령 상태"
        if i==38: return "CDB 인스턴스2 명령 상태"
        if i==39: return "현재 활성 펌웨어 버전"
        if i==85: return "AppDescriptor[0] Host Interface ID"
        if i==86: return "AppDescriptor[0] Media Interface ID"
        if i==87: return "AppDescriptor[0] Host/Media Lane Count"
        if i==88: return "AppDescriptor[0] Host Lane Assign Option"
        if i==126: return "Upper Memory Bank 선택"
        if i==127: return "Upper Memory Page 선택"
        return "N/A"

    def _decoded_p00(self,v,i):
        """Decoded 열: 값을 사람이 읽을 수 있는 형태로"""
        if i==0:  return SFF8024.get(v[0],f"?({h2(v[0])})")
        if i==1:  return asc_display(v,1,16)          # VendorName
        if i==17:
            oui=f"{h2(v[17])}-{h2(v[18])}-{h2(v[19])}"
            # 알려진 OUI
            known={"CC-03-88":"MangoBoost","00-17-6A":"Cisco","00-90-65":"Finisar",
                   "00-02-C9":"Mellanox","00-00-5A":"3Com"}
            name=known.get(oui,"")
            return f"{oui}" + (f" ({name})" if name else "")  # OUI
        if i==20: return asc_display(v,20,35)         # VendorPN
        if i==36: return asc_display(v,36,37)         # VendorRev
        if i==38:
            raw=asc_display(v,38,53)
            # 실제 의미있는 내용 앞부분 추출 (공백 제거)
            stripped=raw.strip()
            if stripped:
                return stripped
            return "(공백 패딩만)"  # VendorSN
        if i==54:
            dc=asc_display(v,54,61)
            if all(v[54+j]==0 for j in range(8)):
                return "(미설정)"
            return dc
        if i==62: return asc_display(v,62,71)         # CLEI
        if i==72:
            pc=(v[72]>>5)&7
            pw=["≤1.5W","≤2.0W","≤2.5W","≤3.5W","≤4.0W","≤4.5W","≤5.0W",">5W"]
            return f"Class{pc+1} {pw[pc]}"
        if i==73: return f"{v[73]*0.25:.2f} W"
        if i==74: return "N/A(cable)" if v[74]==0 else f"Mult={v[74]>>6} Val={v[74]&0x3F}"
        if i==75: return f"Connector={h2(v[75])}"
        if i==84:
            return {0:"850nm VCSEL",1:"1310nm VCSEL",2:"1550nm VCSEL",
                    3:"1310nm FP",4:"1310nm DFB",5:"1490nm DFB",
                    6:"1310nm EML",7:"1550nm EML"}.get(v[84],h2(v[84]))
        if i==94: return f"0x{h2(v[94])}"
        return "N/A"

    def _desc_p00(self,v,i):
        """Description 열: 필드 의미"""
        if i==0:  return "SFF-8024 식별자 사본 (Byte 00h:0 복사)"
        if i==1:  return "제조사 이름 (ASCII 16자, 우측 공백 패딩)"
        if i==17: return "제조사 IEEE OUI (3바이트)"
        if i==20: return "부품번호 (ASCII 16자)"
        if i==36: return "부품번호 리비전 (ASCII 2자)"
        if i==38: return "시리얼번호 (ASCII 16자, 우측 공백 패딩)"
        if i==54: return "제조일 (YYYYMMDD, ASCII 8자)"
        if i==62: return "CLEI 코드 (ASCII 10자, 선택)"
        if i==72: return "전력 등급 (bit7-5) / 최대전력 배수"
        if i==73: return "최대 소비전력 (0.25W 단위)"
        if i==74: return "케이블 길이 정보"
        if i==75: return "커넥터 타입 (SFF-8024 테이블)"
        if i==84: return "미디어 인터페이스 기술"
        if i==94: return "Page 00h 체크섬 (Byte 128~221 합산 mod256)"
        return ""

    # ── P01h (Advertising) ─────────────────────────────
    def _decoded_p01(self,v,i):
        """P01h Decoded: 광고 필드 해석"""
        def s16(hi,lo): return ((hi<<8|lo) if (hi&0x80)==0
                                else -((~(hi<<8|lo)&0xFFFF)+1))
        def u16(hi,lo): return hi<<8|lo
        if i==0:  # InactiveFW Major.Minor
            return f"v{v[0]}.{v[1]}"
        if i==2:  # HW Major.Minor
            return f"HW v{v[2]}.{v[3]}"
        if i==4:  # LengthSMF
            mult=[0.1,1,10,0][v[4]>>6]
            base=v[4]&0x3F
            km=mult*base
            return f"{km:.1f} km" if km else "N/A"
        if i==6:  return f"{v[6]*2} m" if v[6] else "N/A"  # OM5
        if i==7:  return f"{v[7]*2} m" if v[7] else "N/A"  # OM4
        if i==8:  return f"{v[8]*2} m" if v[8] else "N/A"  # OM3
        if i==9:  return f"{v[9]} m"   if v[9] else "N/A"  # OM2
        if i==10: # NominalWavelength (U16 x 0.05nm)
            wl=u16(v[10],v[11])*0.05
            return f"{wl:.2f} nm" if wl else "N/A"
        if i==12: # WavelengthTolerance (U16 x 0.005nm)
            tol=u16(v[12],v[13])*0.005
            return f"±{tol:.3f} nm" if tol else "N/A"
        if i==14: # SupportedPages 비트 플래그
            bits=[]
            if v[14]&0x80: bits.append("P01h")
            if v[14]&0x40: bits.append("P02h")
            if v[14]&0x20: bits.append("P10h")
            if v[14]&0x10: bits.append("P11h")
            return ",".join(bits) if bits else "—"
        if i==127: return f"0x{h2(v[127])}"
        return "—"

    def _desc_p01(self,i):
        """P01h Description"""
        if i==0:  return "비활성 펌웨어 버전 (Major.Minor)"
        if i==2:  return "하드웨어 버전 (Major.Minor)"
        if i==4:  return "SMF 지원 최대 링크 길이 (Byte[7:6]=배율, [5:0]=기본값 km)"
        if i==6:  return "OM5 지원 최대 링크 길이 (단위: 2m)"
        if i==7:  return "OM4 지원 최대 링크 길이 (단위: 2m)"
        if i==8:  return "OM3 지원 최대 링크 길이 (단위: 2m)"
        if i==9:  return "OM2 지원 최대 링크 길이 (단위: 1m)"
        if i==10: return "공칭 파장 (단일 파장 모듈, 단위: 0.05nm)"
        if i==12: return "파장 허용 오차 ±tol (단위: 0.005nm)"
        if i==14: return "지원 페이지 광고 (비트 플래그)"
        if i==15: return "Duration 광고 (State Machine 전환 시간)"
        if i==17: return "모듈 특성 광고 (CDR, TX/RX 기능 등)"
        if i==27: return "지원 Controls 광고"
        if i==29: return "지원 Flags 광고"
        if i==31: return "지원 Monitors 광고"
        if i==33: return "지원 Signal Integrity Controls 광고"
        if i==35: return "지원 CDB 기능 광고"
        if i==127: return "Page 01h 체크섬 (Byte 130~254 합산 mod256)"
        return ""

    # ── P02h (Module and Lane Thresholds) ──────────────
    def _decoded_p02(self,v,i):
        """P02h Decoded: 임계값 실제 단위 변환"""
        def s16(a,b): return ((a<<8|b) if (a&0x80)==0
                              else -((~(a<<8|b)&0xFFFF)+1))
        def u16(a,b): return a<<8|b
        # Temp: S16 / 256 °C (128~135)
        if 0<=i<=6 and i%2==0:
            val=s16(v[i],v[i+1])/256.0
            return f"{val:+.2f} °C"
        # Vcc: U16 × 100µV → V (136~143 → i=8~14)
        if 8<=i<=14 and i%2==0:
            val=u16(v[i],v[i+1])*0.0001
            return f"{val:.4f} V"
        # Aux1/2/3/Custom: S16, 단위 모듈마다 다름 (16~47)
        if 16<=i<=46 and i%2==0:
            val=s16(v[i],v[i+1])
            return f"{val} (raw S16)"
        # Tx Power: U16 × 0.1µW (48~55)
        if 48<=i<=54 and i%2==0:
            val=u16(v[i],v[i+1])*0.1
            return f"{val:.1f} µW"
        # Tx Bias: U16 × 2µA (56~63)
        if 56<=i<=62 and i%2==0:
            val=u16(v[i],v[i+1])*2
            return f"{val} µA"
        # Rx Power: U16 × 0.1µW (64~71)
        if 64<=i<=70 and i%2==0:
            val=u16(v[i],v[i+1])*0.1
            return f"{val:.1f} µW"
        if i==127: return f"0x{h2(v[127])}"
        return "—"

    def _desc_p02(self,i):
        """P02h Description"""
        _temp=["온도 High Alarm","온도 Low Alarm",
               "온도 High Warning","온도 Low Warning"]
        _vcc =["Vcc High Alarm","Vcc Low Alarm",
               "Vcc High Warning","Vcc Low Warning"]
        _aux1=["Aux1 High Alarm","Aux1 Low Alarm",
               "Aux1 High Warning","Aux1 Low Warning"]
        _aux2=["Aux2 High Alarm","Aux2 Low Alarm",
               "Aux2 High Warning","Aux2 Low Warning"]
        _aux3=["Aux3 High Alarm","Aux3 Low Alarm",
               "Aux3 High Warning","Aux3 Low Warning"]
        _cust=["Custom High Alarm","Custom Low Alarm",
               "Custom High Warning","Custom Low Warning"]
        _txpw=["Tx Power High Alarm","Tx Power Low Alarm",
               "Tx Power High Warning","Tx Power Low Warning"]
        _txbi=["Tx Bias High Alarm","Tx Bias Low Alarm",
               "Tx Bias High Warning","Tx Bias Low Warning"]
        _rxpw=["Rx Power High Alarm","Rx Power Low Alarm",
               "Rx Power High Warning","Rx Power Low Warning"]
        table={
            0:(_temp[0],"S16/256 °C"),2:(_temp[1],"S16/256 °C"),
            4:(_temp[2],"S16/256 °C"),6:(_temp[3],"S16/256 °C"),
            8:(_vcc[0], "U16×100µV→V"),10:(_vcc[1],"U16×100µV→V"),
            12:(_vcc[2],"U16×100µV→V"),14:(_vcc[3],"U16×100µV→V"),
            16:(_aux1[0],"S16 TEC전류/Custom"),18:(_aux1[1],"S16"),
            20:(_aux1[2],"S16"),22:(_aux1[3],"S16"),
            24:(_aux2[0],"S16 TEC전류/레이저온도"),26:(_aux2[1],"S16"),
            28:(_aux2[2],"S16"),30:(_aux2[3],"S16"),
            32:(_aux3[0],"S16 레이저온도/추가전압"),34:(_aux3[1],"S16"),
            36:(_aux3[2],"S16"),38:(_aux3[3],"S16"),
            40:(_cust[0],"S16/U16"),42:(_cust[1],"S16/U16"),
            44:(_cust[2],"S16/U16"),46:(_cust[3],"S16/U16"),
            48:(_txpw[0],"U16×0.1µW"),50:(_txpw[1],"U16×0.1µW"),
            52:(_txpw[2],"U16×0.1µW"),54:(_txpw[3],"U16×0.1µW"),
            56:(_txbi[0],"U16×2µA"),58:(_txbi[1],"U16×2µA"),
            60:(_txbi[2],"U16×2µA"),62:(_txbi[3],"U16×2µA"),
            64:(_rxpw[0],"U16×0.1µW"),66:(_rxpw[1],"U16×0.1µW"),
            68:(_rxpw[2],"U16×0.1µW"),70:(_rxpw[3],"U16×0.1µW"),
            127:("Page Checksum","Byte 128~254 합산 mod256"),
        }
        if i in table:
            name,unit=table[i]
            return f"{name} ({unit})"
        return ""

    def _is_dirty(self,pk,addr_hex):
        try:
            s=int(addr_hex.replace("h","").split("~")[0],16)
            if pk=="p00": s-=0x80
            return 0<=s<128 and self.data[pk][s]!=self.orig_data[pk][s]
        except: return False

    def _update_dirty_label(self):
        n=sum(1 for pk in PAGE_KEYS for i in range(128)
              if self.data[pk][i]!=self.orig_data[pk][i])
        self.dirty_label.config(text=f"  ✎ {n}개 수정됨" if n else "",
                                 fg=self.theme["dirty"])

    def _on_inline_edit(self,ri,vals,byte_offset=0):
        """CanvasTable Value[hex] 바이트팝업 commit 콜백
        vals: 바이트 리스트
        byte_offset: 필드 내 시작 오프셋 (단일 바이트 편집 시)
        """
        if ri<0 or ri>=len(self.mct._rows): return
        row=self.mct._rows[ri][0]
        pk=self.current_page
        try:
            base_idx=int(row[1].replace("h","").split("~")[0],16)
            if pk=="p00": base_idx-=0x80
            addr_str=row[1].replace("h","")
            if "~" in addr_str:
                s,e=addr_str.split("~")
                field_len=int(e,16)-int(s,16)+1
            else:
                field_len=1
            write_len=min(len(vals), field_len-byte_offset)
            for i in range(write_len):
                target=base_idx+byte_offset+i
                if 0<=target<128:
                    self.data[pk][target]=vals[i]
            self._show_page(pk)
            val_str=" ".join(h2(v) for v in vals[:write_len])
            self._log(f"편집: {pk} [{row[1]}]+{byte_offset} ← {val_str}")
        except Exception as e:
            logging.error("인라인 편집 오류", exc_info=True)
            messagebox.showerror("오류",str(e))

    def _on_row_dblclick(self,idx):
        vals=self.mct.get_selected_values()
        if not vals: return
        row=vals[0]
        self.edit_info.config(text=f"편집: [{row[1]}]  {row[3]}")
        self.edit_var.set(row[2].split()[0])
        self.edit_entry.focus()
        self._editing_row=row

    def _apply_edit(self):
        row=getattr(self,"_editing_row",None)
        if not row: return
        try: val=int(self.edit_var.get().strip(),16)
        except: messagebox.showwarning("입력 오류","올바른 16진수를 입력하세요."); return
        pk=self.current_page
        try:
            idx=int(row[1].replace("h","").split("~")[0],16)
            if pk=="p00": idx-=0x80
            if 0<=idx<128: self.data[pk][idx]=val&0xFF
        except Exception as e: messagebox.showerror("오류",str(e)); return
        self._show_page(pk); self._log(f"편집: {pk} [{row[1]}] ← {h2(val)}")

    # ── EEPROM 읽기/쓰기 ─────────────────────────────────
    # _read_eeprom : 별도 스레드에서 전체 페이지 읽기
    #   Lower Memory(A0h 00~7F) → 각 Upper Page(80~FF)
    #   페이지 전환: set_page(pn) → Byte 7Fh에 페이지 번호 쓰기
    # _start_write : 선택된 페이지/범위만 쓰기 (dirty_only 옵션)
    # _verify_write: 쓰기 후 read-back 비교 검증
    def _read_eeprom(self):
        if self._device_index < 0 and not self._is_simulated:
            messagebox.showwarning("연결 필요","먼저 연결하세요."); return
        pages_total = 6  # a0 + p00~p11 5개
        def task():
            conn=None
            try:
                conn=self._open_conn()
                i2c=i2c_8bit(self.i2c_addr_var.get())
                self._log("━━━ EEPROM 읽기 시작 ━━━")
                # 상태 표시
                self.after(0,lambda:self.conn_status.config(
                    text="● 읽기 중...", fg=self.theme["amb"]))
                # Lower Memory (A0h)
                self._log(f"[1/6] Lower Memory (A0h) 읽기...")
                data=conn.read_page(i2c, 0x00, 128)
                for reg,val in enumerate(data): self.data["a0"][reg]=val
                self._log(f"  ✓ A0h 완료 ({len(data)}바이트)")
                # Upper Memory Pages
                page_info=list(zip(
                    ["p00","p01","p02","p10","p11"],
                    [0,1,2,0x10,0x11],
                    ["P00h","P01h","P02h","P10h","P11h"]))
                for n,(pk,pn,lbl) in enumerate(page_info,2):
                    self._log(f"[{n}/6] {lbl} (Page 0x{pn:02X}) 읽기...")
                    conn.set_page(pn)
                    data=conn.read_page(i2c, 0x80, 128)
                    for reg,val in enumerate(data): self.data[pk][reg]=val
                    self._log(f"  ✓ {lbl} 완료 ({len(data)}바이트)")
                self.orig_data=copy.deepcopy(self.data)
                self._log("━━━ ✓ 읽기 완료 (6/6 페이지) ━━━")
                self.after(0,lambda:self._show_page(self.current_page))
            except Exception as e:
                logging.error("EEPROM 읽기 오류", exc_info=True)
                self._log(f"읽기 오류: {e} (crash.log 참조)")
                self.after(0,lambda e=e:messagebox.showerror("읽기 오류",str(e)))
            finally:
                self._close_conn(conn)
        threading.Thread(target=task,daemon=True).start()

    def _refresh_summary(self):
        self.write_summary.config(state=tk.NORMAL)
        self.write_summary.delete("1.0",tk.END)
        try: s,e=int(self.write_start.get(),16),int(self.write_end.get(),16)
        except: s,e=0,127
        dirty_only=self.write_dirty_only.get(); total=0
        for pk in PAGE_KEYS:
            if not self.write_vars[pk].get(): continue
            cnt=sum(1 for i in range(s,min(e+1,128))
                    if not dirty_only or self.data[pk][i]!=self.orig_data[pk][i])
            self.write_summary.insert(tk.END,f"{pk.upper():6s}: {cnt:3d}바이트 예정\n")
            total+=cnt
        self.write_summary.insert(tk.END,f"\n합계: {total}바이트")
        self.write_summary.config(state=tk.DISABLED)

    def _start_write(self):
        if not self.conn: messagebox.showwarning("연결 필요","먼저 연결하세요."); return
        if not messagebox.askyesno("쓰기 확인",
            "EEPROM에 데이터를 씁니다.\n되돌릴 수 없습니다. 계속?"): return
        self.write_btn.config(state=tk.DISABLED)
        def task():
            conn=None
            try:
                conn=self._open_conn()
                i2c=i2c_8bit(self.i2c_addr_var.get())  # 7bit→8bit
                s,e=int(self.write_start.get(),16),int(self.write_end.get(),16)
                dirty_only=self.write_dirty_only.get()
                page_map={"p00":0,"p01":1,"p02":2,"p10":0x10,"p11":0x11}
                total=sum(1 for pk in PAGE_KEYS if self.write_vars[pk].get()
                          for i in range(s,min(e+1,128))
                          if not dirty_only or self.data[pk][i]!=self.orig_data[pk][i])
                done=0
                for pk in PAGE_KEYS:
                    if not self.write_vars[pk].get(): continue
                    if pk!="a0": conn.set_page(page_map[pk]); base=0x80
                    else: base=0
                    for i in range(s,min(e+1,128)):
                        if dirty_only and self.data[pk][i]==self.orig_data[pk][i]: continue
                        conn.write_byte(i2c,base+i,self.data[pk][i])
                        done+=1; pct=done/total*100 if total else 100
                        msg=f"{pk.upper()} [{base+i:02X}h]←{h2(self.data[pk][i])} ({done}/{total})"
                        self._log(msg)
                        self.after(0,lambda p=pct,m=msg:(
                            self.progress_var.set(p),self.progress_label.config(text=m)))
                        time.sleep(0.005)
                self.orig_data=copy.deepcopy(self.data); self._log("✓ 쓰기 완료!")
                self.after(0,lambda:(
                    self.progress_label.config(text="✓ 완료"),
                    self._show_page(self.current_page),
                    messagebox.showinfo("완료",f"{done}바이트 쓰기 완료!")))
            except Exception as ex:
                logging.error("EEPROM 쓰기 오류", exc_info=True)
                self._log(f"쓰기 오류: {ex} (crash.log 참조)")
                self.after(0,lambda ex=ex:messagebox.showerror("쓰기 오류",str(ex)))
            finally:
                self._close_conn(conn)
                self.after(0,lambda:self.write_btn.config(state=tk.NORMAL))
        threading.Thread(target=task,daemon=True).start()

    def _verify_write(self):
        if self._device_index < 0 and not self._is_simulated:
            messagebox.showwarning("연결 필요","먼저 연결하세요."); return
        def task():
            conn=None
            try:
                conn=self._open_conn()
                i2c=i2c_8bit(self.i2c_addr_var.get())
                rb=conn.read_page(i2c, 0x00, 128)
                mis=[f"A0[{i:02X}h]: wrote {h2(self.data['a0'][i])}, "
                     f"read {h2(rb[i] if i<len(rb) else None)}"
                     for i in range(128)
                     if (rb[i] if i<len(rb) else -1)!=self.data["a0"][i]]
                if mis:
                    self._log(f"검증 실패: {len(mis)}개")
                    for m in mis[:10]: self._log(f"  {m}")
                    self.after(0,lambda mis=mis:messagebox.showwarning("검증 실패",
                        f"{len(mis)}개 불일치\n"+"\n".join(mis[:5])))
                else:
                    self._log("✓ 검증 통과")
                    self.after(0,lambda:messagebox.showinfo("검증 통과","모든 바이트 일치!"))
            except Exception as ex:
                self._log(f"검증 오류: {ex}")
            finally:
                self._close_conn(conn)
        threading.Thread(target=task,daemon=True).start()

    # ── 로그 ─────────────────────────────────────────────
    def _log(self,msg):
        ts=time.strftime("%H:%M:%S")
        def _ins():
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END,f"[{ts}] {msg}\n")
            self.log_text.see(tk.END); self.log_text.config(state=tk.DISABLED)
        self.after(0,_ins)

    def _clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0",tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _save_log(self):
        path=filedialog.asksaveasfilename(defaultextension=".txt",
                                           filetypes=[("Text","*.txt")])
        if not path: return
        with open(path,"w") as f: f.write(self.log_text.get("1.0",tk.END))


# ── 실행 ──────────────────────────────────────────────────
if __name__=="__main__":
    logging.info(f"=== {APP_NAME} v{APP_VERSION} 시작 ===")
    try:
        app=App()
        app.mainloop()
        logging.info(f"=== 정상 종료 ===")
    except Exception:
        logging.error("치명적 오류로 종료", exc_info=True)
        print(f"오류 발생. 상세: {CRASH_LOG}"); raise
