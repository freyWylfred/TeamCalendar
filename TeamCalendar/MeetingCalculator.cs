namespace TeamCalendar
{
    /// <summary>
    /// 会議時間の計算ロジック（テスト可能なスタティックメソッド）
    /// </summary>
    public static class MeetingCalculator
    {
        /// <summary>
        /// 指定日の会議時間を分単位で算出（勤務時間内にクリップ、休憩除外、重複マージ）
        /// </summary>
        public static double CalcMeetingMinutes(
            IEnumerable<AppointmentInfo> appointments,
            string user,
            DateTime date,
            WorkScheduleConfig config)
        {
            var dayAppts = appointments
                .Where(a => a.Owner == user && a.Start.Date == date && a.ResponseStatus is 3 or 1)
                .OrderBy(a => a.Start)
                .ToList();

            if (dayAppts.Count == 0) return 0;

            var workStart = date.Add(config.StartTime);
            var workEnd = date.Add(config.EndTime);
            var breakStart = date.Add(config.BreakStartTime);
            var breakEnd = date.Add(config.BreakEndTime);

            var intervals = new List<(DateTime s, DateTime e)>();
            foreach (var a in dayAppts)
            {
                var s = a.Start < workStart ? workStart : a.Start;
                var e = a.End > workEnd ? workEnd : a.End;
                if (s >= e) continue;

                if (s < breakStart && e <= breakStart)
                    intervals.Add((s, e));
                else if (s < breakStart && e > breakStart && e <= breakEnd)
                    intervals.Add((s, breakStart));
                else if (s < breakStart && e > breakEnd)
                { intervals.Add((s, breakStart)); intervals.Add((breakEnd, e)); }
                else if (s >= breakStart && s < breakEnd && e <= breakEnd)
                { /* entirely in break */ }
                else if (s >= breakStart && s < breakEnd && e > breakEnd)
                    intervals.Add((breakEnd, e));
                else
                    intervals.Add((s, e));
            }

            if (intervals.Count == 0) return 0;
            intervals = [.. intervals.OrderBy(i => i.s)];

            var merged = new List<(DateTime s, DateTime e)> { intervals[0] };
            for (int i = 1; i < intervals.Count; i++)
            {
                var last = merged[^1];
                if (intervals[i].s <= last.e)
                    merged[^1] = (last.s, intervals[i].e > last.e ? intervals[i].e : last.e);
                else
                    merged.Add(intervals[i]);
            }

            return merged.Sum(m => (m.e - m.s).TotalMinutes);
        }
    }
}
