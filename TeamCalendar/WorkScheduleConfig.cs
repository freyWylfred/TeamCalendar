namespace TeamCalendar
{
    public class WorkScheduleConfig
    {
        public TimeSpan StartTime { get; set; } = new(8, 30, 0);
        public TimeSpan EndTime { get; set; } = new(17, 0, 0);
        public TimeSpan BreakStartTime { get; set; } = new(12, 30, 0);
        public TimeSpan BreakEndTime { get; set; } = new(13, 30, 0);
        public int SlotMinutes { get; set; } = 30;

        public bool IsBreakSlot(TimeSpan slotStart)
        {
            var slotEnd = slotStart.Add(TimeSpan.FromMinutes(SlotMinutes));
            return slotStart < BreakEndTime && slotEnd > BreakStartTime;
        }

        public List<TimeSpan> GenerateTimeSlots()
        {
            var slots = new List<TimeSpan>();
            for (var t = StartTime; t < EndTime; t = t.Add(TimeSpan.FromMinutes(SlotMinutes)))
                slots.Add(t);
            return slots;
        }

        public static WorkScheduleConfig Load(string path)
        {
            var config = new WorkScheduleConfig();
            if (!File.Exists(path)) return config;

            string[] lines;
            try
            {
                lines = File.ReadAllLines(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WARN] 設定ファイルの読み取りに失敗しました ({path}): {ex.Message}");
                return config;
            }

            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (string.IsNullOrEmpty(trimmed) || trimmed[0] is '[' or ';' or '#')
                    continue;

                var eqIndex = trimmed.IndexOf('=');
                if (eqIndex < 0) continue;

                var key = trimmed[..eqIndex].Trim();
                var value = trimmed[(eqIndex + 1)..].Trim();

                switch (key)
                {
                    case "StartTime" when TimeSpan.TryParse(value, out var v): config.StartTime = v; break;
                    case "EndTime" when TimeSpan.TryParse(value, out var v): config.EndTime = v; break;
                    case "BreakStartTime" when TimeSpan.TryParse(value, out var v): config.BreakStartTime = v; break;
                    case "BreakEndTime" when TimeSpan.TryParse(value, out var v): config.BreakEndTime = v; break;
                    case "SlotMinutes" when int.TryParse(value, out var v) && v > 0: config.SlotMinutes = v; break;
                }
            }

            return config;
        }

        public static void CreateDefaultIfMissing(string path)
        {
            if (File.Exists(path)) return;

            try
            {
                File.WriteAllText(path, """
                    [WorkSchedule]
                    ; 勤務開始時刻
                    StartTime=08:30
                    ; 勤務終了時刻
                    EndTime=17:00

                    ; 休憩開始時刻
                    BreakStartTime=12:30
                    ; 休憩終了時刻
                    BreakEndTime=13:30

                    ; タイムラインの時間間隔 (分)
                    SlotMinutes=30
                    """);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WARN] 既定設定ファイルの作成に失敗しました ({path}): {ex.Message}");
            }
        }
    }
}
