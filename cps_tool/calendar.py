"""Working calendar helpers for the CPS calculator."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from typing import Sequence, Tuple


@dataclass(slots=True)
class WorkCalendar:
    """Simple working calendar assuming a constant workday length."""

    workday_start: time = time(hour=8)
    workday_end: time = time(hour=17)
    weekend_days: Tuple[int, ...] = (5, 6)
    hours_per_day: float = field(init=False)

    def __post_init__(self) -> None:
        if self.workday_end <= self.workday_start:
            raise ValueError("workday_end must be after workday_start")
        self.hours_per_day = (
            datetime.combine(date.today(), self.workday_end)
            - datetime.combine(date.today(), self.workday_start)
        ).total_seconds() / 3600.0

    # ------------------------------------------------------------------
    # Alignment helpers

    def align_start(self, moment: datetime) -> datetime:
        """Move *forward* to the next available working time."""

        current = moment
        if current.weekday() in self.weekend_days:
            current = self._next_workday_start(current.date())
        elif current.time() >= self.workday_end:
            current = self._next_workday_start(current.date() + timedelta(days=1))
        elif current.time() < self.workday_start:
            current = datetime.combine(current.date(), self.workday_start)
        return current

    def align_finish(self, moment: datetime) -> datetime:
        """Move *backwards* to the previous available working time."""

        current = moment
        if current.weekday() in self.weekend_days:
            current = self._previous_workday_end(current.date())
        elif current.time() < self.workday_start:
            current = datetime.combine(current.date(), self.workday_start)
        elif current.time() > self.workday_end:
            current = datetime.combine(current.date(), self.workday_end)
        return current

    # ------------------------------------------------------------------
    # Duration arithmetic

    def add_work_duration(self, start: datetime, duration_days: float) -> datetime:
        """Advance a start datetime by the specified working duration."""

        if abs(duration_days) < 1e-9:
            return self.align_start(start)
        hours = duration_days * self.hours_per_day
        return self.add_work_hours(start, hours)

    def add_work_hours(self, start: datetime, hours: float) -> datetime:
        if abs(hours) < 1e-9:
            return self.align_start(start)
        if hours < 0:
            return self.subtract_work_hours(start, -hours)
        current = self.align_start(start)
        remaining = hours
        while remaining > 1e-9:
            day_end = datetime.combine(current.date(), self.workday_end)
            available = (day_end - current).total_seconds() / 3600.0
            if remaining <= available + 1e-9:
                return current + timedelta(hours=remaining)
            remaining -= available
            current = self._next_workday_start(current.date() + timedelta(days=1))
        return current

    def subtract_work_duration(self, finish: datetime, duration_days: float) -> datetime:
        if abs(duration_days) < 1e-9:
            return self.align_finish(finish)
        hours = duration_days * self.hours_per_day
        return self.subtract_work_hours(finish, hours)

    def subtract_work_hours(self, finish: datetime, hours: float) -> datetime:
        if abs(hours) < 1e-9:
            return self.align_finish(finish)
        if hours < 0:
            return self.add_work_hours(finish, -hours)
        current = self.align_finish(finish)
        remaining = hours
        while remaining > 1e-9:
            day_start = datetime.combine(current.date(), self.workday_start)
            available = (current - day_start).total_seconds() / 3600.0
            if remaining <= available + 1e-9:
                return current - timedelta(hours=remaining)
            remaining -= available
            current = self._previous_workday_end(current.date() - timedelta(days=1))
        return current

    def work_hours_between(self, start: datetime, finish: datetime) -> float:
        """Return the number of working hours between two datetimes."""

        if finish <= start:
            return 0.0
        current = self.align_start(start)
        end = self.align_finish(finish)
        if end <= current:
            return 0.0
        if current.date() == end.date():
            return (end - current).total_seconds() / 3600.0

        hours = 0.0
        day_end = datetime.combine(current.date(), self.workday_end)
        hours += (day_end - current).total_seconds() / 3600.0

        next_day = current.date() + timedelta(days=1)
        full_days = self._count_workdays_between(next_day, end.date())
        hours += full_days * self.hours_per_day

        day_start = datetime.combine(end.date(), self.workday_start)
        hours += (end - day_start).total_seconds() / 3600.0
        return hours

    # ------------------------------------------------------------------
    # Helpers

    def _next_workday_start(self, candidate: date) -> datetime:
        next_day = candidate
        while next_day.weekday() in self.weekend_days:
            next_day += timedelta(days=1)
        return datetime.combine(next_day, self.workday_start)

    def _previous_workday_end(self, candidate: date) -> datetime:
        previous_day = candidate
        while previous_day.weekday() in self.weekend_days:
            previous_day -= timedelta(days=1)
        return datetime.combine(previous_day, self.workday_end)

    def _count_workdays_between(self, start_date: date, end_date: date) -> int:
        """Count working days in the half-open interval [start_date, end_date)."""

        if start_date >= end_date:
            return 0
        weekend = set(self.weekend_days)
        total_days = (end_date - start_date).days
        full_weeks, remainder = divmod(total_days, 7)
        workdays = full_weeks * (7 - len(weekend))
        day = start_date + timedelta(days=full_weeks * 7)
        for _ in range(remainder):
            if day.weekday() not in weekend:
                workdays += 1
            day += timedelta(days=1)
        return workdays

    def describe(self) -> str:
        weekend = ", ".join([self._weekday_name(d) for d in self.weekend_days])
        return (
            f"Workday {self.workday_start.strftime('%H:%M')} - {self.workday_end.strftime('%H:%M')} "
            f"({self.hours_per_day:g} hours), weekend: {weekend or 'none'}"
        )

    @staticmethod
    def _weekday_name(index: int) -> str:
        names = [
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
            "Sunday",
        ]
        if 0 <= index < len(names):
            return names[index]
        return str(index)


def parse_weekend(argument: Sequence[str]) -> Tuple[int, ...]:
    """Parse CLI weekend arguments (0=Mon ... 6=Sun)."""

    if not argument:
        return (5, 6)
    indices = []
    for value in argument:
        if value.strip().isdigit():
            indices.append(int(value))
            continue
        normalized = value.strip().lower()
        mapping = {
            "mon": 0,
            "monday": 0,
            "tue": 1,
            "tuesday": 1,
            "wed": 2,
            "wednesday": 2,
            "thu": 3,
            "thursday": 3,
            "fri": 4,
            "friday": 4,
            "sat": 5,
            "saturday": 5,
            "sun": 6,
            "sunday": 6,
        }
        if normalized not in mapping:
            raise ValueError(f"Unable to parse weekend day: {value}")
        indices.append(mapping[normalized])
    return tuple(sorted(set(indices)))
