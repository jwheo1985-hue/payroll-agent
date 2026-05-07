#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
유틸리티 함수 모듈
엑셀 파일 로드 및 전처리
"""

import os
import pandas as pd
import sys


def load_attendance(filepath: str) -> pd.DataFrame:
    """
    S1(세콤) 출퇴근 기록 엑셀 파일 로드
    
    예상 컬럼: 사번, 이름, 일자, 시간
    (파일 형식에 따라 컬럼명이 다를 수 있음)
    """
    if not os.path.exists(filepath):
        print(f"[오류] 파일 없음: {filepath}")
        print("  data/s1_attendance.xlsx 파일을 준비해주세요.")
        print("  S1 출퇴근기록 엑셀을 data/ 폴더에 저장 후 실행하세요.")
        sys.exit(1)

    try:
        df = pd.read_excel(filepath, dtype=str)
        df = df.dropna(how="all")  # 빈 행 제거
        print(f"  [로드] {filepath}: {len(df)}행")
        return df
    except Exception as e:
        print(f"[오류] {filepath} 로드 실패: {e}")
        sys.exit(1)


def load_employees(filepath: str) -> pd.DataFrame:
    """
    ECOUNT 사원 정보 엑셀 파일 로드
    
    예상 컬럼: 사번, 이름, 부서, 직책, 입사일, 퇴사일(선택)
    """
    if not os.path.exists(filepath):
        print(f"[오류] 파일 없음: {filepath}")
        print("  ECOUNT > 인사 > 사원정보 > 엑셀 다운로드 후 data/ 폴더에 저장하세요.")
        sys.exit(1)

    try:
        df = pd.read_excel(filepath, dtype=str)
        df = df.dropna(how="all")
        df.columns = [c.strip() for c in df.columns]
        print(f"  [로드] {filepath}: {len(df)}명")
        return df
    except Exception as e:
        print(f"[오류] {filepath} 로드 실패: {e}")
        sys.exit(1)


def load_hourly_rates(filepath: str) -> pd.DataFrame:
    """
    시급 기준표 엑셀 파일 로드
    
    예상 컬럼: 사번, 이름, 시급
    """
    if not os.path.exists(filepath):
        print(f"[오류] 파일 없음: {filepath}")
        print("  시급 기준표를 data/hourly_rates.xlsx 에 저장하세요.")
        print("  컬럼: 사번 | 이름 | 시급")
        sys.exit(1)

    try:
        df = pd.read_excel(filepath, dtype={"사번": str})
        df = df.dropna(how="all")
        df.columns = [c.strip() for c in df.columns]
        df["시급"] = pd.to_numeric(df["시급"], errors="coerce")
        print(f"  [로드] {filepath}: {len(df)}명")
        return df
    except Exception as e:
        print(f"[오류] {filepath} 로드 실패: {e}")
        sys.exit(1)


def validate_required_files():
    """필수 파일 존재 여부 사전 확인"""
    required = [
        ("data/s1_attendance.xlsx", "S1 출퇴근 기록"),
        ("data/ecount_employees.xlsx", "ECOUNT 사원 정보"),
        ("data/hourly_rates.xlsx", "시급 기준표"),
    ]
    missing = []
    for path, name in required:
        if not os.path.exists(path):
            missing.append(f"  - {path} ({name})")

    if missing:
        print("[필수 파일 누락]")
        for m in missing:
            print(m)
        print()
        print("data/ 폴더에 다음 파일을 준비하세요:")
        print("  1. s1_attendance.xlsx   : S1 출퇴근기록 엑셀")
        print("     컬럼: 사번 | 이름 | 일자 | 시간")
        print("  2. ecount_employees.xlsx: ECOUNT 사원정보 엑셀")
        print("     컬럼: 사번 | 이름 | 부서 | 입사일 | 퇴사일")
        print("  3. hourly_rates.xlsx    : 시급 기준표")
        print("     컬럼: 사번 | 이름 | 시급")
        return False
    return True


def format_currency(value) -> str:
    """금액을 한국식 천단위 구분으로 포맷"""
    try:
        return f"{int(value):,}"
    except (ValueError, TypeError):
        return str(value)
