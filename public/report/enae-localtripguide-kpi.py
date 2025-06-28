import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import os
import re
import math
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx2pdf import convert

# ──────────────────────────────────────────────
# 설정: 컬럼 매핑, 국적 통합, 권역 매핑
# ──────────────────────────────────────────────
COLUMN_ALIASES = {
    '국적': ['국적', '국가'],
    '성별': ['성별'],
    '성명(한글)': ['성명(한글)', '성명(한국어)', '이름'],
    '성명(영문)': ['성명(영문)', '성명(영어)', '이름(영문)', '영문이름'],
    '휴대폰번호': ['휴대폰번호', '연락처', '전화번호', '핸드폰번호']
}

NATIONALITY_MAP = {
    '호주': ['호주', '오스트레일리아'],
    '튀르키예': ['튀르키예', '터키'],
    '우즈베키스탄': ['우즈베키스탄', '우즈벡키스탄', '우주베키스탄', '우즈벡'],
    '카자흐스탄': ['카자흐스탄', '카자흐'],
    '대만': ['대만', '타이완'],
}

REGION_MAP = {
    '동남아시아': ['베트남', '미얀마', '캄보디아', '인도네시아', '필리핀', '라오스', '말레이시아', '태국'],
    '중앙아시아': ['우즈베키스탄', '키르기스스탄', '카자흐스탄', '투르크메니스탄'],
    '동아시아': ['몽골', '일본', '중국', '대만'],
    '남아시아': ['네팔', '방글라데시', '스리랑카', '인도', '파키스탄'],
    '북아메리카': ['미국', '캐나다'],
    '유럽': ['독일', '프랑스', '스웨덴', '이탈리아', '루마니아', '핀란드', '덴마크', '우크라이나',
           '네덜란드', '노르웨이', '스위스', '영국', '체코', '러시아'],
    '아프리카': ['코트디부아르', '나이지리아', '남아프리카공화국', '차드', '콩고', '가봉', '니제르',
             '모로코', '브루나이', '짐바브웨', '탄자니아'],
    '중동': ['튀르키예', '아제르바이잔', '이란', '예멘'],
    '오세아니아': ['호주'],
    '남아메리카': ['브라질', '콜롬비아', '아르헨티나', '페루'],
}

# ──────────────────────────────────────────────
# 유틸리티 함수
# ──────────────────────────────────────────────
def normalize_text(text):
    return str(text).replace(" ", "").lower()

def unify_nationality(nat_value):
    if not isinstance(nat_value, str):
        return nat_value
    nat_value_norm = nat_value.strip().lower()
    for standard, aliases in NATIONALITY_MAP.items():
        if nat_value_norm in [a.lower() for a in aliases]:
            return standard
    return nat_value.strip()

def map_region(nation):
    for region, nations in REGION_MAP.items():
        if nation.lower() in [n.lower() for n in nations]:
            return region
    return '기타'

def truncate_to_2nd_decimal(x):
    return math.floor(x * 100) / 100

def extract_semester(filepath):
    match = re.search(r'(20\d{2})[-_.]?(?:\s*)?(\d)학기', filepath)
    return f"{match.group(1)}-{match.group(2)}학기" if match else "미지정"

def extract_school_name(filepath):
    base = os.path.basename(filepath)
    name_part = os.path.splitext(base)[0]
    parts = name_part.split('_')
    return parts[1].strip() if len(parts) > 1 else "미확인학교"

def map_columns(df, column_aliases):
    col_map = {}
    df_cols_norm = {normalize_text(col): col for col in df.columns}
    for target_col, aliases in column_aliases.items():
        for alias in aliases:
            norm_alias = normalize_text(alias)
            if norm_alias in df_cols_norm:
                col_map[target_col] = df_cols_norm[norm_alias]
                break
    return col_map

def find_header_row(df_preview, column_aliases, min_match_count=2):
    for i in range(len(df_preview)):
        row = df_preview.iloc[i].astype(str).map(normalize_text).tolist()
        match_count = sum(
            any(cell in [normalize_text(a) for a in aliases] for cell in row)
            for aliases in column_aliases.values()
        )
        if match_count >= min_match_count:
            return i
    return None

def calculate_entropy(group_df):
    counts = group_df['국적'].value_counts()
    proportions = counts / counts.sum()
    return -np.sum(proportions * np.log2(proportions))

def school_entropy_scores(df):
    return df.groupby('학교').apply(calculate_entropy).reset_index(name='국제화지표')

def adjust_column_width(filepath):
    wb = load_workbook(filepath)
    ws = wb.active
    align = Alignment(horizontal='center', vertical='center')

    for col in ws.columns:
        max_length = max((len(str(cell.value)) if cell.value else 0) for cell in col)
        for cell in col:
            cell.alignment = align
        ws.column_dimensions[col[0].column_letter].width = max_length + 10

    wb.save(filepath)

def load_and_process_sheet(df, filepath):
    header_row = find_header_row(df.head(30), COLUMN_ALIASES)
    if header_row is None:
        return None

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    col_map = map_columns(df, COLUMN_ALIASES)
    df = df[[v for v in col_map.values() if v]]
    df.columns = [k for k, v in col_map.items() if v]

    mask_all_empty = df.isnull() | (df.astype(str).applymap(lambda x: x.strip() == ''))
    df = df.loc[:mask_all_empty.all(axis=1).idxmax() - 1] if mask_all_empty.all(axis=1).any() else df

    if '국적' in df.columns:
        df = df[df['국적'].notna() & ~df['국적'].astype(str).str.strip().isin(['한국', '대한민국', '', '불명', 'X'])]
        df['국적'] = df['국적'].apply(unify_nationality)

    if '성별' in df.columns:
        df['성별'] = df['성별'].astype(str).str.strip().str.upper().replace({'M': '남', 'F': '여'})

    df.insert(1, '학교', extract_school_name(filepath))
    df.insert(0, '학기', extract_semester(filepath))
    return df

def load_and_process_file(filepath):
    try:
        xls = pd.ExcelFile(filepath)
        dfs = [load_and_process_sheet(pd.read_excel(xls, sheet_name=sheet, header=None), filepath)
               for sheet in xls.sheet_names]
        dfs = [df for df in dfs if df is not None and not df.empty]
        return pd.concat(dfs, ignore_index=True) if dfs else None
    except Exception as e:
        print(f"[오류] {filepath}: {e}")
        return None

def merge_files(filepaths):
    merged = pd.concat(
        [load_and_process_file(path) for path in filepaths if load_and_process_file(path) is not None],
        ignore_index=True
    )
    if merged.empty:
        return merged

    no_phone = merged[merged['휴대폰번호'].isnull() | (merged['휴대폰번호'].astype(str).str.strip() == '')]
    with_phone = merged.dropna(subset=['휴대폰번호'])
    with_phone = with_phone.drop_duplicates(subset=['학기', '휴대폰번호'])

    merged = pd.concat([with_phone, no_phone], ignore_index=True)
    merged = merged.drop_duplicates(subset=['학기', '성명(영문)', '학교', '국적'])
    return merged

def convert_docx_to_pdf(docx_path, pdf_path=None):
    docx_path = os.path.abspath(docx_path)
    pdf_path = os.path.abspath(pdf_path or os.path.splitext(docx_path)[0] + ".pdf")
    convert(docx_path)
    os.replace(os.path.splitext(docx_path)[0] + ".pdf", pdf_path)

def select_files():
    filepaths = filedialog.askopenfilenames(title="엑셀 파일 선택", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not filepaths:
        return

    merged_df = merge_files(filepaths)
    if merged_df.empty:
        messagebox.showwarning("경고", "통합할 데이터가 없습니다.")
        return

    entropy_df = school_entropy_scores(merged_df)
    entropy_df['국제화지표'] = entropy_df['국제화지표'].apply(truncate_to_2nd_decimal)
    merged_df = pd.merge(merged_df, entropy_df, on='학교', how='left')

    save_path = os.path.join(os.path.dirname(filepaths[0]), "로컬트립가이드 종합보고서.xlsx")
    try:
        merged_df.to_excel(save_path, index=False, engine='openpyxl')
        adjust_column_width(save_path)
        messagebox.showinfo("완료", f"저장됨:\n{save_path}")
    except Exception as e:
        messagebox.showerror("오류", f"저장 중 오류:\n{e}")

    # 추가 문서 생성 함수 필요 시 여기에 호출
    # generate_kpi_and_docx() 등

def main():
    root = tk.Tk()
    root.title("엑셀 통합 도구")
    tk.Label(root, text="유학생 리스트 → 종합보고서.xlsx", font=("Arial", 14)).pack(pady=10)
    tk.Button(root, text="엑셀 파일 선택 및 통합", command=select_files, width=30, height=2).pack(pady=20)
    root.mainloop()

if __name__ == "__main__":
    main()