import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "WBS"

# 스타일 정의
header_font = Font(bold=True, size=12)
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
header_font_white = Font(bold=True, size=12, color="FFFFFF")
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# 헤더
headers = ["No", "대", "중", "소", "소요시간"]
col_widths = [6, 18, 40, 45, 15]

for col_idx, (header, width) in enumerate(zip(headers, col_widths), 1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    cell.font = header_font_white
    cell.fill = header_fill
    cell.alignment = center_align
    cell.border = thin_border
    ws.column_dimensions[get_column_letter(col_idx)].width = width

# WBS 데이터: (대, 중, [(소, 소요시간), ...])
data = [
    ("1. Windows OS", [
        ("Windows OS 관련 언어 공부하기", [
            ("C 기초 문법 학습", "3일"),
            ("C++ 기초 문법 학습", "4일"),
            ("Win32 API 개요 파악", "1일"),
            ("Win32 API 핵심 함수 실습", "3일"),
            ("Windows 메모리 관리 학습", "2일"),
            ("Windows 프로세스/스레드 프로그래밍", "5일"),
            ("Windows 파일 시스템 및 I/O 실습", "3일"),
            ("Windows 네트워크 프로그래밍 실습", "5일"),
        ]),
        ("Windows OS 관련 서적 찾고 공부하기", [
            ("관련 서적 목록 조사 및 선정", "0.5일"),
            ("'Windows Internals Part 1' 정독", "7일"),
            ("'Windows Internals Part 2' 정독", "7일"),
            ("'Windows System Programming' 서적 학습", "5일"),
            ("핵심 내용 요약 노트 작성", "1일"),
            ("학습 내용 블로그/문서 정리", "1일"),
        ]),
        ("Windows OS 오픈소스 분석하기", [
            ("분석 대상 오픈소스 프로젝트 선정", "0.5일"),
            ("ReactOS 프로젝트 구조 파악", "2일"),
            ("ReactOS 핵심 모듈 소스코드 분석", "7일"),
            ("Sysinternals 도구 사용법 학습", "1일"),
            ("Sysinternals 소스코드 분석", "5일"),
            ("분석 결과 문서화", "1일"),
        ]),
        ("Windows OS 와 관련된 토이프로젝트 하기", [
            ("프로젝트 아이디어 브레인스토밍", "0.5일"),
            ("프로젝트 요구사항 정의", "0.5일"),
            ("프로젝트 설계 (아키텍처)", "1일"),
            ("개발 환경 세팅", "0.5일"),
            ("핵심 기능 구현", "7일"),
            ("부가 기능 구현", "4일"),
            ("단위 테스트 작성 및 실행", "2일"),
            ("통합 테스트 및 디버깅", "3일"),
            ("README 및 프로젝트 문서 작성", "0.5일"),
            ("코드 리뷰 및 회고", "0.5일"),
        ]),
    ]),
    ("2. Linux OS", [
        ("Linux OS 관련 언어 공부하기", [
            ("C 언어 복습 및 심화 학습", "3일"),
            ("POSIX API 개요 파악", "1일"),
            ("POSIX API 핵심 함수 실습", "3일"),
            ("Shell Script 기본 문법 학습", "1일"),
            ("Shell Script 실전 스크립트 작성", "2일"),
            ("Linux 프로세스/스레드 프로그래밍", "5일"),
            ("Linux 파일 시스템 및 I/O 실습", "3일"),
            ("Linux 소켓 프로그래밍 실습", "5일"),
        ]),
        ("Linux OS 관련 서적 찾고 공부하기", [
            ("관련 서적 목록 조사 및 선정", "0.5일"),
            ("'Linux Kernel Development' 정독", "7일"),
            ("'The Linux Programming Interface' 정독", "7일"),
            ("'Linux Command Line' 서적 학습", "3일"),
            ("핵심 내용 요약 노트 작성", "1일"),
            ("학습 내용 블로그/문서 정리", "1일"),
        ]),
        ("Linux OS 오픈소스 분석하기", [
            ("분석 대상 오픈소스 프로젝트 선정", "0.5일"),
            ("Linux Kernel 전체 구조 파악", "2일"),
            ("Linux Kernel 핵심 서브시스템 분석", "7일"),
            ("systemd 소스코드 분석", "5일"),
            ("busybox 소스코드 분석", "4일"),
            ("분석 결과 문서화", "1일"),
        ]),
        ("Linux OS 와 관련된 토이프로젝트 하기", [
            ("프로젝트 아이디어 브레인스토밍", "0.5일"),
            ("프로젝트 요구사항 정의", "0.5일"),
            ("프로젝트 설계 (아키텍처)", "1일"),
            ("개발 환경 세팅", "0.5일"),
            ("핵심 기능 구현", "7일"),
            ("부가 기능 구현", "4일"),
            ("단위 테스트 작성 및 실행", "2일"),
            ("통합 테스트 및 디버깅", "3일"),
            ("README 및 프로젝트 문서 작성", "0.5일"),
            ("코드 리뷰 및 회고", "0.5일"),
        ]),
    ]),
]

row = 2
no = 1
for dae_name, jungs in data:
    dae_start = row
    for jung_name, sos in jungs:
        jung_start = row
        for so_name, duration in sos:
            ws.cell(row=row, column=1, value=no).alignment = center_align
            ws.cell(row=row, column=1).border = thin_border
            ws.cell(row=row, column=4, value=so_name).alignment = left_align
            ws.cell(row=row, column=4).border = thin_border
            ws.cell(row=row, column=5, value=duration).alignment = center_align
            ws.cell(row=row, column=5).border = thin_border
            no += 1
            row += 1

        # 중 카테고리 병합
        ws.merge_cells(start_row=jung_start, start_column=3, end_row=row-1, end_column=3)
        cell = ws.cell(row=jung_start, column=3, value=jung_name)
        cell.alignment = center_align
        cell.border = thin_border
        # 병합 셀 테두리
        for r in range(jung_start, row):
            ws.cell(row=r, column=3).border = thin_border

    # 대 카테고리 병합
    ws.merge_cells(start_row=dae_start, start_column=2, end_row=row-1, end_column=2)
    cell = ws.cell(row=dae_start, column=2, value=dae_name)
    cell.alignment = center_align
    cell.border = thin_border
    cell.font = Font(bold=True, size=11)
    for r in range(dae_start, row):
        ws.cell(row=r, column=2).border = thin_border

    # 2. Linux OS 대 카테고리에 노란색 배경
    if "Linux" in dae_name:
        for r in range(dae_start, row):
            ws.cell(row=r, column=2).fill = yellow_fill

# 행 높이 설정
for r in range(1, row):
    ws.row_dimensions[r].height = 25

output_path = "/home/user/test/WBS_OS_Study.xlsx"
wb.save(output_path)
print(f"WBS 파일 생성 완료: {output_path}")
