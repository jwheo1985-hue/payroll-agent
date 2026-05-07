#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ECOUNT ERP 시급자 급여 자동정산 프로그램
S1(세콤) 출퇴근 기록 기반 근무시간 계산 및 ECOUNT 입력용 엑셀 생성

작성자: jwheo1985-hue
버전: 1.0.0
"""

import sys
from calculator import PayrollCalculator
from report import ReportGenerator
from utils import load_attendance, load_employees, load_hourly_rates


def main():
    print("=" * 60)
    print("  ECOUNT ERP 시급자 급여 자동정산 프로그램 v1.0")
    print("=" * 60)
    print()

    # 1. 데이터 로드
    print("[1/5] 데이터 로딩 중...")
    attendance_df = load_attendance("data/s1_attendance.xlsx")
    employees_df = load_employees("data/ecount_employees.xlsx")
    hourly_rates_df = load_hourly_rates("data/hourly_rates.xlsx")
    print(f"  - 출퇴근 기록: {len(attendance_df)}건")
    print(f"  - 사원 정보: {len(employees_df)}명")
    print(f"  - 시급 정보: {len(hourly_rates_df)}명")
    print()

    # 2. 출퇴근 데이터 정리
    print("[2/5] 출퇴근 기록 정리 중...")
    calculator = PayrollCalculator(attendance_df, employees_df, hourly_rates_df)
    cleaned_df = calculator.clean_attendance()
    print(f"  - 정리된 기록: {len(cleaned_df)}건")
    print()

    # 3. 근무시간 계산
    print("[3/5] 근무시간 계산 중...")
    work_hours_df = calculator.calculate_work_hours()
    print("  - 기본/연장/야간/휴일/주휴 근무시간 계산 완료")
    print()

    # 4. 급여 계산
    print("[4/5] 급여 계산 중...")
    payroll_df = calculator.calculate_payroll(work_hours_df)
    print("  - 수당별 급여 계산 완료")
    print()

    # 5. 보고서 생성
    print("[5/5] 보고서 생성 중...")
    reporter = ReportGenerator(payroll_df, calculator.errors)
    reporter.generate_all()
    print("  - output/ECOUNT_입력용.xlsx 생성 완료")
    print("  - output/급여검증_오류리스트.xlsx 생성 완료")
    print("  - output/급여마감_체크리스트.xlsx 생성 완료")
    print("  - output/부서별_급여집계.xlsx 생성 완료")
    print()

    # 6. 오류 요약
    if calculator.errors:
        print(f"[!] 확인 필요 항목: {len(calculator.errors)}건")
        for i, err in enumerate(calculator.errors[:10], 1):
            print(f"  {i}. {err}")
        if len(calculator.errors) > 10:
            print(f"  ... 외 {len(calculator.errors)-10}건 (오류리스트 파일 참고)")
    else:
        print("[OK] 오류 없음")

    print()
    print("처리 완료. output/ 폴더를 확인하세요.")
    print("=" * 60)


if __name__ == "__main__":
    main()
