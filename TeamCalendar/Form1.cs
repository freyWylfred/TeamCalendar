using System.Data;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Reflection;
using System.Runtime.InteropServices;
using ClosedXML.Excel;

namespace TeamCalendar
{
    public partial class Form1 : Form
    {
        private readonly List<AppointmentInfo> _appointments = [];
        private WorkScheduleConfig _config = new();
        private readonly Dictionary<(int row, int col), AppointmentInfo?> _timelineCells = [];
        private List<DayOfWeekStats> _dayStats = [];

        // サマリーカード値ラベル
        private Label _lblTotalValue = null!;
        private Label _lblAcceptedValue = null!;
        private Label _lblTentativeValue = null!;
        private Label _lblDeclinedValue = null!;

        #region テーマカラー

        private static readonly Color Accent = Color.FromArgb(0, 120, 212);
        private static readonly Color Success = Color.FromArgb(16, 124, 16);
        private static readonly Color Warning = Color.FromArgb(255, 170, 0);
        private static readonly Color Danger = Color.FromArgb(209, 52, 56);
        private static readonly Color Surface = Color.FromArgb(243, 243, 243);
        private static readonly Color TextSecondary = Color.FromArgb(96, 96, 96);

        private static readonly Color RowAccepted = Color.FromArgb(223, 246, 221);
        private static readonly Color RowOrganizer = Color.FromArgb(208, 228, 245);
        private static readonly Color RowTentative = Color.FromArgb(255, 244, 206);
        private static readonly Color RowDeclined = Color.FromArgb(253, 231, 233);

        private static readonly Color GridHeaderBg = Color.FromArgb(246, 248, 250);
        private static readonly Color GridHeaderFg = Color.FromArgb(60, 60, 60);
        private static readonly Color GridSelection = Color.FromArgb(210, 232, 255);
        private static readonly Color BreakBg = Color.FromArgb(235, 235, 232);

        #endregion

        public Form1()
        {
            InitializeComponent();
            ApplyModernTheme();

            var configDir = Path.GetDirectoryName(Environment.ProcessPath) ?? AppContext.BaseDirectory;
            var configPath = Path.Combine(configDir, "config.ini");
            WorkScheduleConfig.CreateDefaultIfMissing(configPath);
            _config = WorkScheduleConfig.Load(configPath);
            Log($"設定読込: 勤務 {_config.StartTime:hh\\:mm}-{_config.EndTime:hh\\:mm}, 休憩 {_config.BreakStartTime:hh\\:mm}-{_config.BreakEndTime:hh\\:mm}, 間隔 {_config.SlotMinutes}分");

            dtpStart.Value = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + 1);
            dtpEnd.Value = dtpStart.Value.AddDays(4);
            txtUserEmails.PlaceholderText = "例: user1@example.com; user2@example.com";
        }

        #region テーマ適用

        private void ApplyModernTheme()
        {
            // ツールバー下線
            pnlToolbar.Paint += (s, e) =>
            {
                using var pen = new Pen(Color.FromArgb(230, 230, 230), 1);
                e.Graphics.DrawLine(pen, 0, pnlToolbar.Height - 1, pnlToolbar.Width, pnlToolbar.Height - 1);
            };

            // DataGridView ヘッダースタイル
            var headerStyle = dgvAppointments.ColumnHeadersDefaultCellStyle;
            headerStyle.BackColor = GridHeaderBg;
            headerStyle.ForeColor = GridHeaderFg;
            headerStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            headerStyle.SelectionBackColor = GridHeaderBg;
            headerStyle.SelectionForeColor = GridHeaderFg;
            headerStyle.Padding = new Padding(8, 0, 0, 0);

            // DataGridView セルスタイル
            var cellStyle = dgvAppointments.DefaultCellStyle;
            cellStyle.Font = new Font("Segoe UI", 9F);
            cellStyle.ForeColor = Color.FromArgb(30, 30, 30);
            cellStyle.SelectionBackColor = GridSelection;
            cellStyle.SelectionForeColor = Color.FromArgb(30, 30, 30);
            cellStyle.Padding = new Padding(8, 0, 0, 0);

            dgvAppointments.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 252);

            // サマリーカード生成
            InitializeSummaryCards();

            // グラフパネルの設定
            typeof(Panel).InvokeMember("DoubleBuffered",
                BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.NonPublic,
                null, pnlChart, [true]);
            pnlChart.Paint += PaintChart;
        }

        private void InitializeSummaryCards()
        {
            var cards = new (string title, string icon, Color color, string fieldName)[]
            {
                ("全予定", "📋", Accent, "total"),
                ("承認 / 主催", "✅", Success, "accepted"),
                ("任意 (仮)", "⏳", Warning, "tentative"),
                ("辞退", "❌", Danger, "declined"),
            };

            foreach (var (title, icon, color, fieldName) in cards)
            {
                var card = CreateSummaryCard(title, icon, color, out var valueLabel);
                pnlSummary.Controls.Add(card);

                switch (fieldName)
                {
                    case "total": _lblTotalValue = valueLabel; break;
                    case "accepted": _lblAcceptedValue = valueLabel; break;
                    case "tentative": _lblTentativeValue = valueLabel; break;
                    case "declined": _lblDeclinedValue = valueLabel; break;
                }
            }
        }

        private static Panel CreateSummaryCard(string title, string icon, Color accentColor, out Label valueLabel)
        {
            var card = new Panel
            {
                Size = new Size(180, 70),
                Margin = new Padding(6, 0, 6, 0),
                BackColor = Color.White,
            };

            // 左のアクセントバーを描画
            card.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using var brush = new SolidBrush(accentColor);
                g.FillRectangle(brush, 0, 8, 3, card.Height - 16);
            };

            var lblTitle = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 8F),
                ForeColor = Color.FromArgb(110, 110, 110),
                Location = new Point(14, 8),
                Text = $"{icon}  {title}",
            };

            valueLabel = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = accentColor,
                Location = new Point(14, 26),
                Text = "—",
            };

            card.Controls.Add(lblTitle);
            card.Controls.Add(valueLabel);
            return card;
        }

        private void UpdateSummaryCards(int total, int accepted, int tentative, int declined)
        {
            _lblTotalValue.Text = $"{total}";
            _lblAcceptedValue.Text = $"{accepted}";
            _lblTentativeValue.Text = $"{tentative}";
            _lblDeclinedValue.Text = $"{declined}";
        }

        #endregion

        #region イベントハンドラ

        private void btnLoad_Click(object sender, EventArgs e)
        {
            LoadCalendarData();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            ExportAcceptedToExcel();
        }

        private void chkDebugLog_CheckedChanged(object sender, EventArgs e)
        {
            splitMain.Panel2Collapsed = !chkDebugLog.Checked;
        }

        #endregion

        #region ログ

        private void Log(string message)
        {
            string line = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
            Debug.WriteLine(line);
            txtLog.AppendText(line + Environment.NewLine);
        }

        private void LogError(string message, Exception ex)
        {
            Log($"[ERROR] {message}");
            Log($"  例外型: {ex.GetType().FullName}");
            Log($"  メッセージ: {ex.Message}");
            if (ex is COMException comEx)
            {
                Log($"  HRESULT: 0x{comEx.HResult:X8} ({comEx.ErrorCode})");
            }
            if (ex.InnerException is not null)
            {
                Log($"  内部例外: {ex.InnerException.GetType().Name} - {ex.InnerException.Message}");
            }
            Log($"  スタックトレース: {ex.StackTrace}");
        }

        #endregion

        #region Outlook予定取得

        private void LoadCalendarData()
        {
            txtLog.Clear();
            _appointments.Clear();
            _dayStats.Clear();
            dgvAppointments.DataSource = null;
            pnlChart.Invalidate();

            Log("=== Outlook予定取得 開始 ===");

            // 日付バリデーション
            if (dtpStart.Value.Date > dtpEnd.Value.Date)
            {
                Log("[WARN] 開始日が終了日より後です。");
                MessageBox.Show("開始日は終了日以前に設定してください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int maxDays = 90;
            if ((dtpEnd.Value.Date - dtpStart.Value.Date).TotalDays > maxDays)
            {
                Log($"[WARN] 取得期間が{maxDays}日を超えています。");
                var result = MessageBox.Show(
                    $"取得期間が{maxDays}日を超えています。処理に時間がかかる可能性があります。\n続行しますか？",
                    "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result != DialogResult.Yes) return;
            }

            var userEmails = ParseUserEmails();
            if (!chkIncludeSelf.Checked && userEmails.Count == 0)
            {
                Log("[WARN] 対象ユーザーが指定されていません。");
                MessageBox.Show("対象ユーザーのメールアドレスを入力するか、\n「自分の予定を含める」にチェックを入れてください。",
                    "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Log($"取得期間: {dtpStart.Value:yyyy/MM/dd} ～ {dtpEnd.Value:yyyy/MM/dd}");
            Log($"自分の予定: {(chkIncludeSelf.Checked ? "含める" : "含めない")}");
            if (userEmails.Count > 0)
                Log($"他のユーザー: {string.Join("; ", userEmails)}");

            dynamic? outlookApp = null;
            dynamic? ns = null;

            try
            {
                Cursor = Cursors.WaitCursor;
                btnLoad.Enabled = false;
                btnExport.Enabled = false;
                lblStatus.Text = "Outlookから予定を取得中...";
                Application.DoEvents();

                // Outlook COMオブジェクトの生成
                Log("Outlook.Application の ProgID を取得中...");
                var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType is null)
                {
                    Log("[ERROR] Outlook.Application の ProgID が見つかりません。Outlookがインストールされていない可能性があります。");
                    MessageBox.Show(
                        "Outlookがインストールされていません。\n\nMicrosoft Outlookをインストールしてから再度お試しください。",
                        "Outlook未検出", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                Log($"ProgID 取得成功: {outlookType.FullName}");

                Log("Outlook.Application インスタンスを作成中...");
                outlookApp = Activator.CreateInstance(outlookType)
                    ?? throw new InvalidOperationException("Outlook.Application のインスタンス作成に失敗しました。");
                Log("Outlook.Application インスタンス作成成功");

                Log("MAPI 名前空間を取得中...");
                ns = outlookApp.GetNamespace("MAPI")
                    ?? throw new InvalidOperationException("MAPI 名前空間の取得に失敗しました。");
                Log("MAPI 名前空間 取得成功");

                string startDate = dtpStart.Value.Date.ToString("yyyy/MM/dd HH:mm");
                string endDate = dtpEnd.Value.Date.AddDays(1).ToString("yyyy/MM/dd HH:mm");
                string filter = $"[Start] >= '{startDate}' AND [Start] < '{endDate}'";
                Log($"フィルター: {filter}");

                int totalProcessed = 0;
                int totalSkipped = 0;

                // 自分の予定を取得
                if (chkIncludeSelf.Checked)
                {
                    Log("--- 自分の予定表を取得中 ---");
                    dynamic? selfFolder = null;
                    try
                    {
                        selfFolder = ns.GetDefaultFolder(9) // olFolderCalendar
                            ?? throw new InvalidOperationException("自分の予定表フォルダの取得に失敗しました。");
                        Log("自分の予定表フォルダ 取得成功");
                        (int p, int s) = ReadCalendarFromFolder((object)selfFolder!, "自分", filter);
                        totalProcessed += p;
                        totalSkipped += s;
                    }
                    finally
                    {
                        SafeReleaseCom(selfFolder, "selfFolder");
                    }
                }

                // 他のユーザーの予定を取得
                foreach (string email in userEmails)
                {
                    Log($"--- {email} の予定表を取得中 ---");
                    dynamic? recipient = null;
                    dynamic? sharedFolder = null;
                    try
                    {
                        recipient = ns.CreateRecipient(email);
                        recipient.Resolve();
                        if (!(bool)recipient.Resolved)
                        {
                            Log($"[WARN] ユーザー '{email}' を解決できませんでした。Exchange上に存在しないか、メールアドレスが正しくありません。");
                            continue;
                        }
                        Log($"ユーザー '{email}' の解決に成功");

                        sharedFolder = ns.GetSharedDefaultFolder(recipient, 9); // olFolderCalendar
                        Log($"'{email}' の共有予定表フォルダ 取得成功");
                        (int p, int s) = ReadCalendarFromFolder((object)sharedFolder!, email, filter);
                        totalProcessed += p;
                        totalSkipped += s;
                    }
                    catch (COMException comEx)
                    {
                        Log($"[ERROR] '{email}' の予定表にアクセスできません: HRESULT=0x{comEx.HResult:X8} {comEx.Message}");
                        Log("  ※ 相手の予定表が共有されているか、アクセス権限があるか確認してください。");
                    }
                    catch (Exception ex)
                    {
                        Log($"[ERROR] '{email}' の予定取得中にエラー: {ex.GetType().Name}: {ex.Message}");
                    }
                    finally
                    {
                        SafeReleaseCom(sharedFolder, $"folder({email})");
                        SafeReleaseCom(recipient, $"recipient({email})");
                    }
                }

                Log($"全ユーザー列挙完了: 処理={totalProcessed}件, 取得={_appointments.Count}件, スキップ={totalSkipped}件");

                if (totalSkipped > 0)
                {
                    Log($"[WARN] {totalSkipped}件の予定が読み取れずスキップされました。");
                }

                BindTimelineGrid();
                CalculateDayStats();
                pnlChart.Invalidate();

                int totalCount = _appointments.Count;
                int acceptedCount = _appointments.Count(a => a.ResponseStatus is 3 or 1);
                int tentativeCount = _appointments.Count(a => a.ResponseStatus == 2);
                int declinedCount = _appointments.Count(a => a.ResponseStatus == 4);

                UpdateSummaryCards(totalCount, acceptedCount, tentativeCount, declinedCount);

                string statusText = $"取得完了: 全{totalCount}件 (承認/主催: {acceptedCount}件, 任意: {tentativeCount}件, 辞退: {declinedCount}件)";
                if (totalSkipped > 0)
                {
                    statusText += $" ※スキップ: {totalSkipped}件";
                }
                lblStatus.Text = statusText;
                Log($"=== 完了: {statusText} ===");
            }
            catch (COMException ex) when (ex.HResult == unchecked((int)0x80080005))
            {
                LogError("サーバー実行に失敗しました (CO_E_SERVER_EXEC_FAILURE)", ex);
                MessageBox.Show(
                    "Outlookの起動に失敗しました。\n\n" +
                    "以下を確認してください:\n" +
                    "・Outlookが正しくインストールされているか\n" +
                    "・別のダイアログ（パスワード入力等）が表示されていないか\n" +
                    "・管理者権限で実行が必要ではないか\n\n" +
                    $"HRESULT: 0x{ex.HResult:X8}\n{ex.Message}",
                    "Outlook起動エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "エラー: Outlookの起動に失敗しました";
            }
            catch (COMException ex) when (ex.HResult == unchecked((int)0x80040154))
            {
                LogError("COMクラスが登録されていません (REGDB_E_CLASSNOTREG)", ex);
                MessageBox.Show(
                    "OutlookのCOMコンポーネントが登録されていません。\n\n" +
                    "Outlookの修復インストールを実行してください。\n\n" +
                    $"HRESULT: 0x{ex.HResult:X8}\n{ex.Message}",
                    "COM登録エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "エラー: Outlook COM未登録";
            }
            catch (COMException ex)
            {
                LogError("Outlook COM通信エラー", ex);
                MessageBox.Show(
                    "Outlookとの通信中にエラーが発生しました。\n\n" +
                    "以下を確認してください:\n" +
                    "・Outlookが起動しているか\n" +
                    "・Outlookがフリーズしていないか\n" +
                    "・Outlookのプロファイルが正しく設定されているか\n\n" +
                    $"HRESULT: 0x{ex.HResult:X8}\n{ex.Message}",
                    "Outlook通信エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "エラー: Outlook通信失敗";
            }
            catch (InvalidOperationException ex)
            {
                LogError("Outlookオブジェクトの初期化エラー", ex);
                MessageBox.Show(
                    $"Outlookの初期化に失敗しました。\n\n{ex.Message}",
                    "初期化エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "エラー: 初期化失敗";
            }
            catch (UnauthorizedAccessException ex)
            {
                LogError("アクセス権限エラー", ex);
                MessageBox.Show(
                    "Outlookへのアクセスが拒否されました。\n\n" +
                    "Outlookのセキュリティ設定でプログラムからのアクセスが許可されているか確認してください。",
                    "アクセス拒否", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "エラー: アクセス拒否";
            }
            catch (Exception ex)
            {
                LogError("予期しないエラー", ex);
                MessageBox.Show(
                    $"予期しないエラーが発生しました。\n\n" +
                    $"例外型: {ex.GetType().Name}\n{ex.Message}\n\n" +
                    "詳細はデバッグログを確認してください。",
                    "予期しないエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "エラー: 予期しないエラー";
            }
            finally
            {
                SafeReleaseCom(ns, "ns");
                SafeReleaseCom(outlookApp, "outlookApp");

                btnLoad.Enabled = true;
                btnExport.Enabled = true;
                Cursor = Cursors.Default;
                Log("=== Outlook予定取得 終了 ===");
            }
        }

        private (int processed, int skipped) ReadCalendarFromFolder(
            object folderObj, string ownerName, string filter)
        {
            dynamic folder = folderObj;
            dynamic? items = null;
            dynamic? restrictedItems = null;
            int processedCount = 0;
            int skippedCount = 0;

            try
            {
                items = folder.Items
                    ?? throw new InvalidOperationException($"'{ownerName}' の予定アイテムコレクションを取得できません。");

                items.Sort("[Start]");
                items.IncludeRecurrences = true;
                Log($"  Sort/IncludeRecurrences 設定完了 ({ownerName})");

                restrictedItems = items.Restrict(filter);
                Log($"  Restrict 実行成功 ({ownerName})。予定アイテムの列挙を開始...");

                foreach (dynamic item in restrictedItems)
                {
                    processedCount++;
                    try
                    {
                        int itemClass = SafeGetProperty<int>(item, "Class", -1);
                        if (itemClass != 26) // olAppointment
                        {
                            Log($"    #{processedCount}: Class={itemClass} のためスキップ (olAppointment=26 以外)");
                            skippedCount++;
                            continue;
                        }

                        string subject = SafeGetProperty<string>(item, "Subject", "(件名取得失敗)");
                        DateTime start = SafeGetProperty<DateTime>(item, "Start", DateTime.MinValue);
                        DateTime end = SafeGetProperty<DateTime>(item, "End", DateTime.MinValue);
                        int duration = SafeGetProperty<int>(item, "Duration", 0);
                        string organizer = SafeGetProperty<string>(item, "Organizer", "(取得失敗)");
                        string location = SafeGetProperty<string>(item, "Location", "");
                        int responseStatus = SafeGetProperty<int>(item, "ResponseStatus", -1);

                        string status = GetStatusText(responseStatus);

                        _appointments.Add(new AppointmentInfo
                        {
                            Owner = ownerName,
                            Subject = subject,
                            Start = start,
                            End = end,
                            Duration = duration,
                            Organizer = organizer,
                            Location = location,
                            Status = status,
                            ResponseStatus = responseStatus
                        });

                        Log($"    #{processedCount}: [{status}] {start:MM/dd HH:mm}-{end:HH:mm} {subject}");
                    }
                    catch (COMException comEx)
                    {
                        skippedCount++;
                        Log($"    #{processedCount}: [SKIP] COM例外 HRESULT=0x{comEx.HResult:X8}: {comEx.Message}");
                    }
                    catch (InvalidCastException castEx)
                    {
                        skippedCount++;
                        Log($"    #{processedCount}: [SKIP] 型変換失敗: {castEx.Message}");
                    }
                    catch (Exception ex)
                    {
                        skippedCount++;
                        Log($"    #{processedCount}: [SKIP] {ex.GetType().Name}: {ex.Message}");
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(item);
                    }
                }
            }
            finally
            {
                SafeReleaseCom(restrictedItems, $"restrictedItems({ownerName})");
                SafeReleaseCom(items, $"items({ownerName})");
            }

            Log($"  [{ownerName}] 完了: 処理={processedCount}件, スキップ={skippedCount}件");
            return (processedCount, skippedCount);
        }

        private List<string> ParseUserEmails()
        {
            return [.. txtUserEmails.Text
                .Split([';', ',', '\n', '\r'], StringSplitOptions.RemoveEmptyEntries)
                .Select(e => e.Trim())
                .Where(e => e.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)];
        }

        #endregion

        #region タイムライン表示

        private void BindTimelineGrid()
        {
            Log("タイムライングリッドへのバインドを開始...");
            _timelineCells.Clear();

            try
            {
                var dt = new DataTable();
                dt.Columns.Add("ユーザー", typeof(string));
                dt.Columns.Add("日付", typeof(string));

                var slots = _config.GenerateTimeSlots();
                foreach (var slot in slots)
                    dt.Columns.Add(slot.ToString(@"hh\:mm"), typeof(string));

                // ユーザーリスト構築（入力順を維持）
                var users = new List<string>();
                if (chkIncludeSelf.Checked) users.Add("自分");
                foreach (var email in ParseUserEmails())
                {
                    if (!users.Contains(email, StringComparer.OrdinalIgnoreCase))
                        users.Add(email);
                }

                var startDate = dtpStart.Value.Date;
                var endDate = dtpEnd.Value.Date;

                int rowIdx = 0;
                foreach (var user in users)
                {
                    for (var date = startDate; date <= endDate; date = date.AddDays(1))
                    {
                        if (date.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday)
                            continue;

                        var row = dt.NewRow();
                        row["ユーザー"] = user;
                        row["日付"] = date.ToString("MM/dd(ddd)");

                        var dayAppts = _appointments
                            .Where(a => a.Owner == user && a.Start.Date == date)
                            .OrderBy(a => a.Start)
                            .ToList();

                        var shownAppts = new HashSet<AppointmentInfo>();

                        for (int si = 0; si < slots.Count; si++)
                        {
                            var slotStart = date.Add(slots[si]);
                            var slotEnd = slotStart.AddMinutes(_config.SlotMinutes);
                            int col = si + 2;

                            var slotAppts = dayAppts.Where(a => a.Start < slotEnd && a.End > slotStart).ToList();
                            var apt = slotAppts.FirstOrDefault();
                            _timelineCells[(rowIdx, col)] = apt;

                            if (slotAppts.Count > 1)
                            {
                                row[col] = $"{slotAppts.Count}件";
                                foreach (var sa in slotAppts) shownAppts.Add(sa);
                            }
                            else if (apt != null && !shownAppts.Contains(apt))
                            {
                                row[col] = TruncateText(apt.Subject, 8);
                                shownAppts.Add(apt);
                            }
                            else
                            {
                                row[col] = "";
                            }
                        }

                        dt.Rows.Add(row);
                        rowIdx++;
                    }
                }

                dgvAppointments.DataSource = dt;

                // 列幅設定
                dgvAppointments.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                if (dgvAppointments.Columns.Count > 0)
                {
                    dgvAppointments.Columns[0].Width = 110;
                    dgvAppointments.Columns[0].Frozen = true;
                    dgvAppointments.Columns[1].Width = 88;
                    dgvAppointments.Columns[1].Frozen = true;

                    for (int c = 2; c < dgvAppointments.Columns.Count; c++)
                    {
                        dgvAppointments.Columns[c].Width = 72;
                        dgvAppointments.Columns[c].DefaultCellStyle.Font = new Font("Segoe UI", 7.5F);
                        dgvAppointments.Columns[c].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }
                }

                // セル色・ツールチップ適用
                string prevUser = "";
                bool altUser = false;
                for (int r = 0; r < dgvAppointments.Rows.Count; r++)
                {
                    var currentUser = dgvAppointments.Rows[r].Cells[0].Value?.ToString() ?? "";
                    if (currentUser != prevUser)
                    {
                        altUser = !altUser;
                        prevUser = currentUser;
                    }

                    // ユーザー・日付列の背景（ユーザー交互色）
                    var baseBg = altUser ? Color.FromArgb(250, 250, 252) : Color.White;
                    dgvAppointments.Rows[r].Cells[0].Style.BackColor = baseBg;
                    dgvAppointments.Rows[r].Cells[1].Style.BackColor = baseBg;

                    for (int c = 2; c < dgvAppointments.Columns.Count; c++)
                    {
                        var cell = dgvAppointments.Rows[r].Cells[c];

                        if (_timelineCells.TryGetValue((r, c), out var apt) && apt != null)
                        {
                            cell.Style.BackColor = apt.ResponseStatus switch
                            {
                                3 => RowAccepted,
                                1 => RowOrganizer,
                                2 => RowTentative,
                                4 => RowDeclined,
                                _ => Color.White
                            };
                            cell.ToolTipText = $"{apt.Subject}\n{apt.Start:HH:mm} \u2013 {apt.End:HH:mm}\n{apt.Status} / {apt.Organizer}";
                        }
                        else
                        {
                            int si = c - 2;
                            if (si < slots.Count && _config.IsBreakSlot(slots[si]))
                            {
                                cell.Style.BackColor = BreakBg;
                            }
                        }
                    }
                }

                Log($"タイムライングリッド バインド完了: {dt.Rows.Count}行, {slots.Count}タイムスロット");
            }
            catch (Exception ex)
            {
                LogError("タイムライングリッドのバインドに失敗しました", ex);
            }
        }

        private static string TruncateText(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text.Length <= maxLength ? text : text[..(maxLength - 1)] + "\u2026";
        }

        #endregion

        #region グラフ

        private void CalculateDayStats()
        {
            var workMinutes = (_config.EndTime - _config.StartTime).TotalMinutes
                            - (_config.BreakEndTime - _config.BreakStartTime).TotalMinutes;

            var startDate = dtpStart.Value.Date;
            var endDate = dtpEnd.Value.Date;
            var weekdays = new[] { DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday };

            var users = _appointments.Select(a => a.Owner).Distinct().ToList();
            if (users.Count == 0)
            {
                _dayStats = [];
                return;
            }

            _dayStats = weekdays.Select(dow =>
            {
                var dates = new List<DateTime>();
                for (var d = startDate; d <= endDate; d = d.AddDays(1))
                    if (d.DayOfWeek == dow) dates.Add(d);

                if (dates.Count == 0) return null;

                double totalMeetingMinutes = 0;
                int userDayCount = 0;

                foreach (var user in users)
                {
                    foreach (var date in dates)
                    {
                        totalMeetingMinutes += CalcUserMeetingMinutes(user, date);
                        userDayCount++;
                    }
                }

                double avg = userDayCount > 0 ? totalMeetingMinutes / userDayCount : 0;
                return new DayOfWeekStats(dow, Math.Min(avg, workMinutes), workMinutes);
            })
            .OfType<DayOfWeekStats>()
            .ToList();

            Log($"曜日別統計算出完了: {string.Join(", ", _dayStats.Select(s => $"{DayLabel(s.Day)}={s.MeetingMinutes:F0}分"))}");
        }

        private double CalcUserMeetingMinutes(string user, DateTime date)
        {
            // 承認(3) / 主催者(1) のみグラフ集計対象
            var dayAppts = _appointments
                .Where(a => a.Owner == user && a.Start.Date == date && a.ResponseStatus is 3 or 1)
                .OrderBy(a => a.Start)
                .ToList();

            if (dayAppts.Count == 0) return 0;

            var workStart = date.Add(_config.StartTime);
            var workEnd = date.Add(_config.EndTime);
            var breakStart = date.Add(_config.BreakStartTime);
            var breakEnd = date.Add(_config.BreakEndTime);

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

        private void PaintChart(object? sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            var panel = (Panel)sender!;
            var rect = panel.ClientRectangle;

            // カード風の白背景
            var cardRect = new Rectangle(16, 0, rect.Width - 32, rect.Height - 8);
            using (var bgBrush = new SolidBrush(Color.White))
                g.FillRectangle(bgBrush, cardRect);

            if (_dayStats.Count == 0)
            {
                using var font = new Font("Segoe UI", 9.5F);
                using var brush = new SolidBrush(TextSecondary);
                var msg = "\U0001f4ca  予定を取得するとグラフが表示されます";
                var sz = g.MeasureString(msg, font);
                g.DrawString(msg, font, brush, (rect.Width - sz.Width) / 2, (rect.Height - sz.Height) / 2);
                return;
            }

            int leftMargin = 72;
            int rightMargin = 32;
            int topMargin = 38;
            int bottomMargin = 30;
            int legendHeight = 24;

            var chartLeft = cardRect.Left + leftMargin;
            var chartTop = cardRect.Top + topMargin;
            var chartBottom = cardRect.Bottom - bottomMargin - legendHeight;
            var chartRight = cardRect.Right - rightMargin;
            var chartWidth = chartRight - chartLeft;
            var chartHeight = chartBottom - chartTop;

            if (chartWidth <= 0 || chartHeight <= 0) return;

            // タイトル
            using (var titleFont = new Font("Segoe UI", 10F, FontStyle.Bold))
            using (var titleBrush = new SolidBrush(Color.FromArgb(40, 40, 40)))
                g.DrawString("\U0001f4ca  曜日別  会議時間 / 空き時間", titleFont, titleBrush, cardRect.Left + 14, cardRect.Top + 8);

            // Y軸スケール
            double maxHours = Math.Ceiling(_dayStats.Max(s => s.WorkMinutes) / 60.0);
            if (maxHours < 1) maxHours = 1;

            using var gridPen = new Pen(Color.FromArgb(238, 238, 238), 1);
            using var axisFont = new Font("Segoe UI", 7.5F);
            using var axisBrush = new SolidBrush(Color.FromArgb(130, 130, 130));

            int ySteps = (int)maxHours;
            for (int h = 0; h <= ySteps; h++)
            {
                float y = chartBottom - (float)(h / maxHours * chartHeight);
                g.DrawLine(gridPen, chartLeft, y, chartRight, y);
                var label = $"{h}h";
                var labelSz = g.MeasureString(label, axisFont);
                g.DrawString(label, axisFont, axisBrush, chartLeft - labelSz.Width - 8, y - labelSz.Height / 2);
            }

            // バー描画
            var meetingColor = Color.FromArgb(0, 120, 212);
            var freeColor = Color.FromArgb(195, 230, 195);

            float barAreaWidth = (float)chartWidth / _dayStats.Count;
            float barWidth = Math.Min(barAreaWidth * 0.55f, 80);
            float barGap = (barAreaWidth - barWidth) / 2;

            using var meetBrush = new SolidBrush(meetingColor);
            using var freeBrush = new SolidBrush(freeColor);
            using var barLabelFont = new Font("Segoe UI", 7.5F, FontStyle.Bold);
            using var dayFont = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            using var whiteBrush = new SolidBrush(Color.White);
            using var greenLabelBrush = new SolidBrush(Color.FromArgb(50, 110, 50));
            using var darkBrush = new SolidBrush(Color.FromArgb(60, 60, 60));
            using var pctBrush = new SolidBrush(Color.FromArgb(90, 90, 90));

            for (int i = 0; i < _dayStats.Count; i++)
            {
                var stat = _dayStats[i];
                float x = chartLeft + i * barAreaWidth + barGap;

                double meetH = stat.MeetingMinutes / 60.0;
                double freeH = stat.FreeMinutes / 60.0;

                float meetBarH = (float)(meetH / maxHours * chartHeight);
                float freeBarH = (float)(freeH / maxHours * chartHeight);

                // 空き時間バー（上部）
                if (freeBarH > 1)
                    g.FillRectangle(freeBrush, x, chartBottom - meetBarH - freeBarH, barWidth, freeBarH);

                // 会議時間バー（下部）
                if (meetBarH > 1)
                    g.FillRectangle(meetBrush, x, chartBottom - meetBarH, barWidth, meetBarH);

                // 会議時間ラベル
                if (meetBarH > 18)
                {
                    var ml = $"{meetH:F1}h";
                    var ms = g.MeasureString(ml, barLabelFont);
                    g.DrawString(ml, barLabelFont, whiteBrush,
                        x + (barWidth - ms.Width) / 2,
                        chartBottom - meetBarH + (meetBarH - ms.Height) / 2);
                }

                // 空き時間ラベル
                if (freeBarH > 18)
                {
                    var fl = $"{freeH:F1}h";
                    var fs = g.MeasureString(fl, barLabelFont);
                    g.DrawString(fl, barLabelFont, greenLabelBrush,
                        x + (barWidth - fs.Width) / 2,
                        chartBottom - meetBarH - freeBarH + (freeBarH - fs.Height) / 2);
                }

                // 会議率パーセンテージ（バー上部）
                double pct = stat.WorkMinutes > 0 ? stat.MeetingMinutes / stat.WorkMinutes * 100 : 0;
                var pctLabel = $"{pct:F0}%";
                var pctSz = g.MeasureString(pctLabel, barLabelFont);
                g.DrawString(pctLabel, barLabelFont, pctBrush,
                    x + (barWidth - pctSz.Width) / 2,
                    chartBottom - meetBarH - freeBarH - pctSz.Height - 2);

                // 曜日ラベル
                var dayName = DayLabel(stat.Day);
                var daySz = g.MeasureString(dayName, dayFont);
                g.DrawString(dayName, dayFont, darkBrush,
                    x + (barWidth - daySz.Width) / 2, chartBottom + 4);
            }

            // X軸線
            using var axisPen = new Pen(Color.FromArgb(210, 210, 210), 1);
            g.DrawLine(axisPen, chartLeft, chartBottom, chartRight, chartBottom);

            // 凡例
            float legendY = cardRect.Bottom - legendHeight - 2;
            float legendX = chartLeft;

            using var legendFont = new Font("Segoe UI", 8F);
            using var legendTextBrush = new SolidBrush(Color.FromArgb(90, 90, 90));

            g.FillRectangle(meetBrush, legendX, legendY + 3, 14, 14);
            g.DrawString("会議時間", legendFont, legendTextBrush, legendX + 18, legendY + 2);

            legendX += 95;
            g.FillRectangle(freeBrush, legendX, legendY + 3, 14, 14);
            g.DrawString("空き時間", legendFont, legendTextBrush, legendX + 18, legendY + 2);

            // 勤務時間注釈
            legendX += 95;
            var note = $"（勤務 {_config.StartTime:hh\\:mm}–{_config.EndTime:hh\\:mm}  休憩除く）";
            g.DrawString(note, legendFont, axisBrush, legendX, legendY + 2);
        }

        private static string DayLabel(DayOfWeek dow) => dow switch
        {
            DayOfWeek.Monday => "月",
            DayOfWeek.Tuesday => "火",
            DayOfWeek.Wednesday => "水",
            DayOfWeek.Thursday => "木",
            DayOfWeek.Friday => "金",
            _ => dow.ToString()
        };

        #endregion

        #region Excel出力

        private void ExportAcceptedToExcel()
        {
            Log("=== Excel出力 開始 ===");

            var accepted = _appointments
                .Where(a => a.ResponseStatus is 3 or 1)
                .OrderBy(a => a.Owner)
                .ThenBy(a => a.Start)
                .ToList();

            if (accepted.Count == 0)
            {
                Log("[WARN] 承認済みの会議が0件です。");
                MessageBox.Show("承認済みの会議がありません。\n先に「予定を取得」で予定を読み込んでください。",
                    "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Log($"承認済み会議: {accepted.Count}件");

            using var dialog = new SaveFileDialog
            {
                Filter = "Excelファイル (*.xlsx)|*.xlsx",
                FileName = $"承認済み会議_{dtpStart.Value:yyyyMMdd}_{dtpEnd.Value:yyyyMMdd}.xlsx"
            };

            if (dialog.ShowDialog() != DialogResult.OK)
            {
                Log("ユーザーがキャンセルしました。");
                return;
            }

            string filePath = dialog.FileName;
            Log($"保存先: {filePath}");

            try
            {
                Cursor = Cursors.WaitCursor;
                btnExport.Enabled = false;
                lblStatus.Text = "Excel出力中...";
                Application.DoEvents();

                using var workbook = new XLWorkbook();
                var ws = workbook.Worksheets.Add("承認済み会議");

                string[] headers = ["ユーザー", "件名", "開始日時", "終了日時", "時間(分)", "開催者", "場所", "ステータス"];
                for (int c = 0; c < headers.Length; c++)
                {
                    ws.Cell(1, c + 1).Value = headers[c];
                }

                var headerRange = ws.Range(1, 1, 1, headers.Length);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
                headerRange.Style.Font.FontColor = XLColor.White;

                for (int i = 0; i < accepted.Count; i++)
                {
                    int row = i + 2;
                    ws.Cell(row, 1).Value = accepted[i].Owner;
                    ws.Cell(row, 2).Value = accepted[i].Subject;
                    ws.Cell(row, 3).Value = accepted[i].Start;
                    ws.Cell(row, 4).Value = accepted[i].End;
                    ws.Cell(row, 5).Value = accepted[i].Duration;
                    ws.Cell(row, 6).Value = accepted[i].Organizer;
                    ws.Cell(row, 7).Value = accepted[i].Location;
                    ws.Cell(row, 8).Value = accepted[i].Status;

                    ws.Cell(row, 3).Style.NumberFormat.Format = "yyyy/MM/dd HH:mm";
                    ws.Cell(row, 4).Style.NumberFormat.Format = "yyyy/MM/dd HH:mm";
                }

                ws.Columns().AdjustToContents();
                Log("ワークブック作成完了。ファイルに保存中...");

                workbook.SaveAs(filePath);

                string msg = $"Excel出力完了: {accepted.Count}件の承認済み会議を出力しました";
                lblStatus.Text = msg;
                Log(msg);
                MessageBox.Show($"Excelファイルを保存しました。\n{filePath}",
                    "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (IOException ex) when (ex.HResult == unchecked((int)0x80070020))
            {
                LogError("ファイルが別のプロセスで使用中です", ex);
                MessageBox.Show(
                    $"ファイルが別のプログラムで開かれています。\n\n" +
                    $"ファイルを閉じてから再度お試しください。\n{filePath}",
                    "ファイルロック", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (IOException ex)
            {
                LogError("ファイルI/Oエラー", ex);
                MessageBox.Show(
                    $"ファイルの書き込み中にエラーが発生しました。\n\n" +
                    $"保存先: {filePath}\n{ex.Message}",
                    "I/Oエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (UnauthorizedAccessException ex)
            {
                LogError("ファイル書き込み権限エラー", ex);
                MessageBox.Show(
                    $"ファイルへの書き込み権限がありません。\n\n" +
                    $"保存先: {filePath}\n\n" +
                    "別のフォルダに保存するか、管理者権限で実行してください。",
                    "アクセス拒否", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                LogError("Excel出力エラー", ex);
                MessageBox.Show(
                    $"Excel出力に失敗しました。\n\n" +
                    $"例外型: {ex.GetType().Name}\n{ex.Message}\n\n" +
                    "詳細はデバッグログを確認してください。",
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnExport.Enabled = true;
                Cursor = Cursors.Default;
                Log("=== Excel出力 終了 ===");
            }
        }

        #endregion

        #region ユーティリティ

        private static T SafeGetProperty<T>(dynamic comObject, string propertyName, T fallback)
        {
            try
            {
                object? val = ((object)comObject).GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty,
                    null,
                    comObject,
                    null);
                return val is T typed ? typed : fallback;
            }
            catch
            {
                return fallback;
            }
        }

        private void SafeReleaseCom(dynamic? comObject, string name)
        {
            if (comObject is null) return;
            try
            {
                Marshal.ReleaseComObject(comObject);
                Log($"COM解放: {name}");
            }
            catch (Exception ex)
            {
                Log($"[WARN] COM解放失敗 ({name}): {ex.Message}");
            }
        }

        private static string GetStatusText(int responseStatus) => responseStatus switch
        {
            0 => "未設定",
            1 => "主催者",
            2 => "任意",
            3 => "承認",
            4 => "辞退",
            5 => "未応答",
            _ => $"不明({responseStatus})"
        };

        #endregion
    }

    public class AppointmentInfo
    {
        public string Owner { get; set; } = "";
        public string Subject { get; set; } = "";
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public int Duration { get; set; }
        public string Organizer { get; set; } = "";
        public string Location { get; set; } = "";
        public string Status { get; set; } = "";
        public int ResponseStatus { get; set; }
    }

    public record DayOfWeekStats(DayOfWeek Day, double MeetingMinutes, double WorkMinutes)
    {
        public double FreeMinutes => Math.Max(0, WorkMinutes - MeetingMinutes);
    }
}
