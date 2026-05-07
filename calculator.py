#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
시급자 급여 계산 모듈
S1 출퇴근 기록 기반 근무시간 및 수당 계산
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time
from typing import List, Dict, Tuple

# 상수 정의
NIGHT_START = time(22, 0)   # 야간근로 시작 (22:00)
NIGHT_END = time(6, 0)      # 야간근로 종료 (06:00)
DAILY_BASIC_HOURS = 8       # 1일 기본근로시간
WEEKLY_HOURS_FOR_JUJU = 15  # 주휴수당 발생 최소 주간근로시간

# 수당 배율
EXTENDED_RATE = 1.5         # 연장근로 (기본 1.0 + 가산 0.5)
NIGHT_EXTRA_RATE = 0.5      # 야간가산 (기본임금에 0.5 추가)
HOLIDAY_RATE = 1.5          # 휴일근로 8시간 이내 (기본 1.0 + 가산 0.5)
HOLIDAY_EXTENDED_RATE = 2.0 # 휴일근로 8시간 초과

# 공휴일 목록 (2026년)
PUBLIC_HOLIDAYS_2026 = [
    "2026-01-01", "2026-01-28", "2026-01-29", "2026-01-30",
    "2026-03-01", "2026-05-05", "2026-05-25", "2026-06-06",
    "2026-08-15", "2026-09-24", "2026-09-25", "2026-09-26",
    "2026-10-03", "2026-10-09", "2026-12-25"
]


class PayrollCalculator:
    def __init__(self, attendance_df, employees_df, hourly_rates_df):
        self.attendance_df = attendance_df.copy()
        self.employees_df = employees_df.copy()
        self.hourly_rates_df = hourly_rates_df.copy()
        self.errors = []
        self.holidays = [pd.Timestamp(d) for d in PUBLIC_HOLIDAYS_2026]

    def is_holiday(self, date) -> bool:
        """해당 날짜가 공휴일 또는 일요일인지 확인"""
        if isinstance(date, str):
            date = pd.Timestamp(date)
        return date.weekday() == 6 or date in self.holidays

    def is_saturday(self, date) -> bool:
        """토요일 여부 확인"""
        if isinstance(date, str):
            date = pd.Timestamp(date)
        return date.weekday() == 5

    def clean_attendance(self) -> pd.DataFrame:
        """S1 출퇴근 기록 정리: 사번별 일자별 최초 출근/최종 퇴근 추출"""
        df = self.attendance_df.copy()

        # 컬럼명 정규화
        df.columns = [c.strip() for c in df.columns]
        if "사번" not in df.columns:
            df = df.rename(columns={df.columns[0]: "사번", df.columns[1]: "이름",
                                    df.columns[2]: "일자", df.columns[3]: "시간"})

        df["일자"] = pd.to_datetime(df["일자"]).dt.date
        df["시간"] = pd.to_datetime(df["시간"], format="%H:%M", errors="coerce").dt.time

        # 사번이 ECOUNT 사원정보에 없는 경우 오류 등록
        valid_ids = set(self.employees_df["사번"].astype(str).unique())
        for emp_id in df["사번"].astype(str).unique():
            if emp_id not in valid_ids:
                self.errors.append(f"사번 불일치: {emp_id} (S1 기록에 있으나 ECOUNT에 없음)")

        # 퇴사자 근무기록 확인
        if "퇴사일" in self.employees_df.columns:
            for _, emp in self.employees_df.dropna(subset=["퇴사일"]).iterrows():
                emp_records = df[df["사번"].astype(str) == str(emp["사번"])]
                if len(emp_records) > 0:
                    last_date = emp_records["일자"].max()
                    quit_date = pd.Timestamp(emp["퇴사일"]).date()
                    if last_date > quit_date:
                        self.errors.append(f"퇴사자 근무기록: {emp.get('이름', emp['사번'])} ({emp['퇴사일']} 이후 기록 존재)")

        # 일자별 최초 출근 / 최종 퇴근
        result = df.groupby(["사번", "이름", "일자"]).agg(
            출근=("시간", "min"),
            퇴근=("시간", "max")
        ).reset_index()

        # 출근만 있고 퇴근 없는 경우 (동일 시간 = 출근=퇴근)
        mask = result["출근"] == result["퇴근"]
        for _, row in result[mask].iterrows():
            self.errors.append(f"출퇴근 기록 이상: {row.get('이름', row['사번'])} {row['일자']} (단일 기록만 존재)")

        return result

    def calc_night_hours(self, start: datetime, end: datetime) -> float:
        """야간근로시간(22:00~06:00) 계산 (시간 단위)"""
        night_hours = 0.0
        current = start

        while current < end:
            next_hour = current + timedelta(hours=1)
            if next_hour > end:
                next_hour = end
            h = current.hour
            # 22시~23시, 0시~5시 = 야간
            if h >= 22 or h < 6:
                night_hours += (next_hour - current).total_seconds() / 3600
            current = next_hour

        return round(night_hours, 2)

    def calculate_work_hours(self) -> pd.DataFrame:
        """일자별 근무유형별 시간 계산"""
        cleaned = self.clean_attendance()
        records = []

        for _, row in cleaned.iterrows():
            emp_id = str(row["사번"])
            name = row.get("이름", emp_id)
            date = row["일자"]
            checkin = row["출근"]
            checkout = row["퇴근"]

            if pd.isna(checkin) or pd.isna(checkout):
                self.errors.append(f"시간 데이터 오류: {name} {date}")
                continue

            # datetime 변환
            dt_date = datetime.combine(date, checkin)
            # 퇴근이 출근보다 이른 경우(날짜 넘김)
            if checkout < checkin:
                dt_checkout = datetime.combine(date + timedelta(days=1), checkout)
            else:
                dt_checkout = datetime.combine(date, checkout)

            total_hours = (dt_checkout - dt_date).total_seconds() / 3600

            # 휴게시간 차감 (근로기준법 기준)
            if total_hours >= 8:
                break_hours = 1.0
            elif total_hours >= 4:
                break_hours = 0.5
            else:
                break_hours = 0.0

            work_hours = total_hours - break_hours

            # 12시간 초과 경보
            if work_hours > 12:
                self.errors.append(f"12시간 초과 근무: {name} {date} ({work_hours:.1f}시간)")

            # 휴일 여부 판단
            is_hol = self.is_holiday(date)
            is_sat = self.is_saturday(date)

            # 근무시간 분류
            if is_hol:
                # 휴일근로
                holiday_basic = min(work_hours, 8.0)
                holiday_extended = max(work_hours - 8.0, 0.0)
                basic_hours = 0.0
                extended_hours = 0.0
            elif is_sat:
                # 토요일 (약정휴일 여부에 따라 다름, 기본 연장으로 처리)
                basic_hours = 0.0
                extended_hours = work_hours
                holiday_basic = 0.0
                holiday_extended = 0.0
            else:
                # 평일
                basic_hours = min(work_hours, DAILY_BASIC_HOURS)
                extended_hours = max(work_hours - DAILY_BASIC_HOURS, 0.0)
                holiday_basic = 0.0
                holiday_extended = 0.0

            # 야간근로시간
            night_hours = self.calc_night_hours(dt_date, dt_checkout)

            records.append({
                "사번": emp_id,
                "이름": name,
                "일자": date,
                "총체류시간": round(total_hours, 2),
                "휴게시간": break_hours,
                "인정근무시간": round(work_hours, 2),
                "기본근로시간": round(basic_hours, 2),
                "연장근로시간": round(extended_hours, 2),
                "야간근로시간": round(night_hours, 2),
                "휴일기본시간": round(holiday_basic, 2),
                "휴일연장시간": round(holiday_extended, 2),
                "휴일여부": is_hol,
            })

        return pd.DataFrame(records)

    def calculate_weekly_juju(self, work_df: pd.DataFrame) -> pd.DataFrame:
        """주휴수당 계산"""
        result = []
        work_df["주"] = pd.to_datetime(work_df["일자"]).dt.isocalendar().week.astype(int)
        work_df["연도"] = pd.to_datetime(work_df["일자"]).dt.isocalendar().year.astype(int)

        for (emp_id, year, week), grp in work_df.groupby(["사번", "연도", "주"]):
            total_week_hours = grp["기본근로시간"].sum() + grp["연장근로시간"].sum()
            has_juju = total_week_hours >= WEEKLY_HOURS_FOR_JUJU
            juju_hours = 8.0 if has_juju else 0.0
            if has_juju:
                result.append({"사번": emp_id, "연도": year, "주": week, "주휴시간": juju_hours})

        return pd.DataFrame(result) if result else pd.DataFrame(columns=["사번", "연도", "주", "주휴시간"])

    def calculate_payroll(self, work_df: pd.DataFrame) -> pd.DataFrame:
        """월별 사원별 수당 계산"""
        if work_df.empty:
            return pd.DataFrame()

        # 주휴수당 계산
        juju_df = self.calculate_weekly_juju(work_df)
        juju_by_emp = juju_df.groupby("사번")["주휴시간"].sum().reset_index()

        # 월별 집계
        monthly = work_df.groupby("사번").agg(
            이름=("이름", "first"),
            근무일수=("일자", "nunique"),
            기본근로시간=("기본근로시간", "sum"),
            연장근로시간=("연장근로시간", "sum"),
            야간근로시간=("야간근로시간", "sum"),
            휴일기본시간=("휴일기본시간", "sum"),
            휴일연장시간=("휴일연장시간", "sum"),
        ).reset_index()

        # 주휴시간 합산
        monthly = monthly.merge(juju_by_emp, on="사번", how="left")
        monthly["주휴시간"] = monthly["주휴시간"].fillna(0)

        # 시급 정보 병합
        hourly_map = dict(zip(
            self.hourly_rates_df["사번"].astype(str),
            self.hourly_rates_df["시급"]
        ))

        monthly["시급"] = monthly["사번"].astype(str).map(hourly_map)

        # 시급 미등록 오류
        missing_rate = monthly[monthly["시급"].isna()]
        for _, row in missing_rate.iterrows():
            self.errors.append(f"시급 미등록: {row.get('이름', row['사번'])}")
        monthly["시급"] = monthly["시급"].fillna(0)

        # 수당 계산
        monthly["기본급"] = (monthly["기본근로시간"] * monthly["시급"]).round()
        monthly["연장근로수당"] = (monthly["연장근로시간"] * monthly["시급"] * EXTENDED_RATE).round()
        monthly["야간근로수당"] = (monthly["야간근로시간"] * monthly["시급"] * NIGHT_EXTRA_RATE).round()
        monthly["휴일근로수당"] = (
            monthly["휴일기본시간"] * monthly["시급"] * HOLIDAY_RATE +
            monthly["휴일연장시간"] * monthly["시급"] * HOLIDAY_EXTENDED_RATE
        ).round()
        monthly["주휴수당"] = (monthly["주휴시간"] * monthly["시급"]).round()
        monthly["총지급액"] = (
            monthly["기본급"] + monthly["연장근로수당"] +
            monthly["야간근로수당"] + monthly["휴일근로수당"] + monthly["주휴수당"]
        )

        return monthly
