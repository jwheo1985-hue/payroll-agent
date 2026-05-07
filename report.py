#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
보고서 생성 모듈
ECOUNT 입력용 엑셀, 오류리스트, 마감 체크리스트, 부서별 집계 생성
"""

import os
import pandas as pd
from datetime import datetime
from typing import List


class ReportGenerator:
    def __init__(self, payroll_df: pd.DataFrame, errors: List[str]):
        self.payroll_df = payroll_df
        self.errors = errors
        self.output_dir = "output"
        os.makedirs(self.output_dir, exist_ok=True)
        self.today = datetime.today().strftime("%Y%m%d")

    def generate_all(self):
        """모든 보고서 생성"""
        self.generate_ecount_input()
        self.generate_error_list()
        self.generate_checklist()
        self.generate_dept_summary()

    def generate_ecount_input(self):
        """ECOUNT 급여 입력용 엑셀 생성"""
        if self.payroll_df.empty:
            return

        cols = [
            "사번", "이름", "시급",
            "기본근로시간", "연장근로시간", "야간근로시간",
            "휴일기본시간", "휴일연장시간", "주휴시간",
            "기본급", "연장근로수당", "야간근로수당",
            "휴일근로수당", "주휴수당", "총지급액"
        ]
        available = [c for c in cols if c in self.payroll_df.columns]
        df = self.payroll_df[available].copy()

        # 숫자 컬럼 정수 변환
        for col in ["기본급", "연장근로수당", "야간근로수당", "휴일근로수당", "주휴수당", "총지급액"]:
            if col in df.columns:
                df[col] = df[col].fillna(0).astype(int)

        filepath = os.path.join(self.output_dir, f"ECOUNT_입력용_{self.today}.xlsx")
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="ECOUNT입력용")

            ws = writer.sheets["ECOUNT입력용"]
            # 열 너비 자동 조정
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 4

        print(f"  저장: {filepath}")

    def generate_error_list(self):
        """오류 및 확인 필요 항목 리스트 생성"""
        filepath = os.path.join(self.output_dir, f"급여검증_오류리스트_{self.today}.xlsx")

        if not self.errors:
            error_df = pd.DataFrame({"오류내용": ["오류 없음"]})
        else:
            error_df = pd.DataFrame({
                "번호": range(1, len(self.errors) + 1),
                "오류내용": self.errors,
                "확인여부": [""] * len(self.errors),
                "처리결과": [""] * len(self.errors),
            })

        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            error_df.to_excel(writer, index=False, sheet_name="오류리스트")
            ws = writer.sheets["오류리스트"]
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

        print(f"  저장: {filepath}")

    def generate_checklist(self):
        """월 급여 마감 전 체크리스트 생성"""
        filepath = os.path.join(self.output_dir, f"급여마감_체크리스트_{self.today}.xlsx")

        checklist = [
            {"항목": "S1 출퇴근 데이터 최종 확인", "완료": ""},
            {"항목": "ECOUNT 사원정보 업데이트 확인", "완료": ""},
            {"항목": "신규 입사자 사번/시급 등록 확인", "완료": ""},
            {"항목": "퇴사자 급여 일할계산 확인", "완료": ""},
            {"항목": "오류리스트 전체 처리 확인", "완료": ""},
            {"항목": "야간/연장/휴일수당 계산 검토", "완료": ""},
            {"항목": "주휴수당 대상자 최종 확인", "완료": ""},
            {"항목": "ECOUNT 급여항목 입력 완료", "완료": ""},
            {"항목": "급여대장 결재 완료", "완료": ""},
            {"항목": "급여 이체 처리", "완료": ""},
            {"항목": "급여명세서(UserPay) 발행", "완료": ""},
            {"항목": "4대보험 취득/상실 신고 대상 확인", "완료": ""},
        ]

        check_df = pd.DataFrame(checklist)
        check_df["비고"] = ""

        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            check_df.to_excel(writer, index=False, sheet_name="마감체크리스트")
            ws = writer.sheets["마감체크리스트"]
            ws.column_dimensions["A"].width = 40
            ws.column_dimensions["B"].width = 10
            ws.column_dimensions["C"].width = 30

        print(f"  저장: {filepath}")

    def generate_dept_summary(self):
        """부서별 급여 집계 생성"""
        if self.payroll_df.empty:
            return

        filepath = os.path.join(self.output_dir, f"부서별_급여집계_{self.today}.xlsx")

        # 부서 정보가 없는 경우 전체 합계만
        summary_cols = ["기본급", "연장근로수당", "야간근로수당", "휴일근로수당", "주휴수당", "총지급액"]
        available = [c for c in summary_cols if c in self.payroll_df.columns]

        total = self.payroll_df[available].sum().to_frame("합계").T
        total.insert(0, "구분", "전체합계")
        total.insert(1, "인원수", len(self.payroll_df))

        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            # 전체 합계
            total.to_excel(writer, index=False, sheet_name="전체집계")
            # 개인별 명세
            self.payroll_df.to_excel(writer, index=False, sheet_name="개인별명세")

            for sheet_name in writer.sheets:
                ws = writer.sheets[sheet_name]
                for col in ws.columns:
                    max_len = max(len(str(cell.value or "")) for cell in col)
                    ws.column_dimensions[col[0].column_letter].width = max_len + 4

        print(f"  저장: {filepath}")
