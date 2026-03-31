namespace TeamCalendar
{
    /// <summary>
    /// Exchange Server 接続情報を入力するダイアログ
    /// </summary>
    public sealed class ExchangeLoginDialog : Form
    {
        private readonly TextBox _txtServerUrl;
        private readonly TextBox _txtEmail;
        private readonly TextBox _txtPassword;
        private readonly TextBox _txtDomain;
        private readonly CheckBox _chkIgnoreSsl;
        private readonly Button _btnTest;
        private readonly Button _btnConnect;
        private readonly Button _btnDisconnect;
        private readonly Button _btnCancel;
        private readonly Label _lblTestResult;

        /// <summary>入力された認証情報（接続時に設定）</summary>
        public ExchangeCredential? Credential { get; private set; }

        /// <summary>切断が選択された場合 true</summary>
        public bool Disconnected { get; private set; }

        // ── Theme colors ──
        private static readonly Color Accent = Color.FromArgb(99, 102, 241);
        private static readonly Color AccentHover = Color.FromArgb(79, 70, 229);
        private static readonly Color AccentPressed = Color.FromArgb(67, 56, 202);
        private static readonly Color Success = Color.FromArgb(5, 150, 105);
        private static readonly Color Danger = Color.FromArgb(220, 38, 38);
        private static readonly Color HeaderBg = Color.FromArgb(30, 41, 59);
        private static readonly Color TextSecondary = Color.FromArgb(100, 116, 139);
        private static readonly Color BorderColor = Color.FromArgb(226, 232, 240);

        public ExchangeLoginDialog(ExchangeCredential? existing = null)
        {
            Text = "Exchange Server 接続設定";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(460, 500);
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9F);

            // ── Header ──
            var pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 48,
                BackColor = HeaderBg,
            };
            var lblTitle = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(16, 12),
                Text = "🔗  Exchange Server 接続設定",
            };
            pnlHeader.Controls.Add(lblTitle);
            Controls.Add(pnlHeader);

            int y = 64;
            const int labelX = 24;
            const int inputX = 24;
            const int inputW = 410;

            // ── EWS URL ──
            Controls.Add(CreateLabel("EWS エンドポイント URL", labelX, y));
            y += 20;
            _txtServerUrl = CreateTextBox(inputX, y, inputW);
            _txtServerUrl.PlaceholderText = "https://mail.example.com/EWS/Exchange.asmx";
            Controls.Add(_txtServerUrl);
            y += 34;

            // ── Email ──
            Controls.Add(CreateLabel("メールアドレス", labelX, y));
            y += 20;
            _txtEmail = CreateTextBox(inputX, y, inputW);
            _txtEmail.PlaceholderText = "user@example.com";
            Controls.Add(_txtEmail);
            y += 34;

            // ── Password ──
            Controls.Add(CreateLabel("パスワード", labelX, y));
            y += 20;
            _txtPassword = CreateTextBox(inputX, y, inputW);
            _txtPassword.UseSystemPasswordChar = true;
            Controls.Add(_txtPassword);
            y += 34;

            // ── Domain ──
            Controls.Add(CreateLabel("ドメイン（NTLM 認証時のみ・任意）", labelX, y));
            y += 20;
            _txtDomain = CreateTextBox(inputX, y, inputW);
            _txtDomain.PlaceholderText = "DOMAIN";
            Controls.Add(_txtDomain);
            y += 38;

            // ── Ignore SSL ──
            _chkIgnoreSsl = new CheckBox
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = TextSecondary,
                Location = new Point(inputX, y),
                Text = "⚠️ SSL 証明書エラーを無視する（自己署名証明書向け）",
            };
            Controls.Add(_chkIgnoreSsl);
            y += 32;

            // ── Test button ──
            _btnTest = new Button
            {
                Text = "⚡ 接続テスト",
                Location = new Point(inputX, y),
                Size = new Size(130, 34),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(241, 245, 249),
                ForeColor = Color.FromArgb(51, 65, 85),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand,
            };
            _btnTest.FlatAppearance.BorderColor = BorderColor;
            _btnTest.Click += BtnTest_Click;
            Controls.Add(_btnTest);

            // ── Test result ──
            _lblTestResult = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = TextSecondary,
                Location = new Point(inputX + 140, y + 8),
                Text = "",
            };
            Controls.Add(_lblTestResult);
            y += 50;

            // ── Buttons ──
            _btnConnect = new Button
            {
                Text = "🔗 接続",
                Size = new Size(110, 36),
                Location = new Point(460 - 110 - 120 - 24 - 12, y),
                FlatStyle = FlatStyle.Flat,
                BackColor = Accent,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand,
            };
            _btnConnect.FlatAppearance.BorderSize = 0;
            _btnConnect.FlatAppearance.MouseOverBackColor = AccentHover;
            _btnConnect.FlatAppearance.MouseDownBackColor = AccentPressed;
            _btnConnect.Click += BtnConnect_Click;
            Controls.Add(_btnConnect);

            _btnCancel = new Button
            {
                Text = "キャンセル",
                Size = new Size(110, 36),
                Location = new Point(460 - 110 - 24, y),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(241, 245, 249),
                ForeColor = Color.FromArgb(51, 65, 85),
                Font = new Font("Segoe UI", 9F),
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Cancel,
            };
            _btnCancel.FlatAppearance.BorderColor = BorderColor;
            Controls.Add(_btnCancel);
            y += 46;

            _btnDisconnect = new Button
            {
                Text = "🔌 切断（Outlook COM に戻す）",
                Size = new Size(240, 34),
                Location = new Point(inputX, y),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Danger,
                Font = new Font("Segoe UI", 8.5F),
                Cursor = Cursors.Hand,
                Visible = existing?.IsConfigured == true,
            };
            _btnDisconnect.FlatAppearance.BorderColor = Color.FromArgb(254, 202, 202);
            _btnDisconnect.Click += BtnDisconnect_Click;
            Controls.Add(_btnDisconnect);

            CancelButton = _btnCancel;

            // ── Pre-populate ──
            if (existing?.IsConfigured == true)
            {
                _txtServerUrl.Text = existing.ServerUrl;
                _txtEmail.Text = existing.Email;
                _txtPassword.Text = existing.Password;
                _txtDomain.Text = existing.Domain;
                _chkIgnoreSsl.Checked = existing.IgnoreSslErrors;
            }
        }

        private async void BtnTest_Click(object? sender, EventArgs e)
        {
            if (!ValidateInputs()) return;

            _btnTest.Enabled = false;
            _lblTestResult.ForeColor = TextSecondary;
            _lblTestResult.Text = "接続テスト中...";
            Application.DoEvents();

            try
            {
                using var svc = new ExchangeCalendarService(BuildCredential());
                await svc.TestConnectionAsync();

                _lblTestResult.ForeColor = Success;
                _lblTestResult.Text = "✅ 接続テスト成功";
            }
            catch (Exception ex)
            {
                _lblTestResult.ForeColor = Danger;
                _lblTestResult.Text = $"❌ 失敗: {Truncate(ex.Message, 50)}";

                MessageBox.Show(
                    $"接続テストに失敗しました。\n\n{ex.Message}",
                    "接続テスト失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _btnTest.Enabled = true;
            }
        }

        private void BtnConnect_Click(object? sender, EventArgs e)
        {
            if (!ValidateInputs()) return;

            Credential = BuildCredential();
            Disconnected = false;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnDisconnect_Click(object? sender, EventArgs e)
        {
            Credential = null;
            Disconnected = true;
            DialogResult = DialogResult.OK;
            Close();
        }

        private bool ValidateInputs()
        {
            if (string.IsNullOrWhiteSpace(_txtServerUrl.Text))
            {
                MessageBox.Show("EWS エンドポイント URL を入力してください。",
                    "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtServerUrl.Focus();
                return false;
            }

            if (!Uri.TryCreate(_txtServerUrl.Text.Trim(), UriKind.Absolute, out var uri)
                || (uri.Scheme != "https" && uri.Scheme != "http"))
            {
                MessageBox.Show("有効な URL を入力してください。\n例: https://mail.example.com/EWS/Exchange.asmx",
                    "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtServerUrl.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(_txtEmail.Text))
            {
                MessageBox.Show("メールアドレスを入力してください。",
                    "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtEmail.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(_txtPassword.Text))
            {
                MessageBox.Show("パスワードを入力してください。",
                    "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtPassword.Focus();
                return false;
            }

            return true;
        }

        private ExchangeCredential BuildCredential() => new()
        {
            ServerUrl = _txtServerUrl.Text.Trim(),
            Email = _txtEmail.Text.Trim(),
            Password = _txtPassword.Text,
            Domain = _txtDomain.Text.Trim(),
            IgnoreSslErrors = _chkIgnoreSsl.Checked,
        };

        private static Label CreateLabel(string text, int x, int y) => new()
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 8.5F),
            ForeColor = Color.FromArgb(100, 116, 139),
            Location = new Point(x, y),
            Text = text,
        };

        private static TextBox CreateTextBox(int x, int y, int width) => new()
        {
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Segoe UI", 9.5F),
            Location = new Point(x, y),
            Size = new Size(width, 25),
        };

        private static string Truncate(string text, int max) =>
            text.Length > max ? text[..max] + "..." : text;
    }
}
