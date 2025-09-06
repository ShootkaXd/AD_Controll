using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AD
{
    // ------------------------- Плоский TabControl (заголовки только) -------------------------
    internal class FlatTabControl : TabControl
    {
        public bool DarkMode { get; set; } = false;

        public FlatTabControl()
        {
            SetStyle(ControlStyles.UserPaint |
                     ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer, true);

            DrawMode = TabDrawMode.OwnerDrawFixed;
            SizeMode = TabSizeMode.Fixed;
            ItemSize = new Size(140, 28);
            Padding = new Point(12, 4);
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            var sel = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            var rect = e.Bounds;
            var page = TabPages[e.Index];

            Color back = sel
                ? (DarkMode ? Color.FromArgb(43, 49, 58) : Color.White)
                : (DarkMode ? Color.FromArgb(34, 40, 49) : Color.FromArgb(246, 248, 250));
            Color text = DarkMode ? Color.Gainsboro : Color.FromArgb(33, 37, 41);

            using (var b = new SolidBrush(back)) e.Graphics.FillRectangle(b, rect);
            TextRenderer.DrawText(e.Graphics, page.Text, Font, rect, text,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

            if (!DarkMode && sel)
            {
                using var p = new Pen(Color.FromArgb(28, 164, 97), 2);
                e.Graphics.DrawLine(p, rect.Left, rect.Bottom - 1, rect.Right, rect.Bottom - 1);
            }
        }
    }

    // -------------------------------------------- Форма --------------------------------------------
    public class MainForm : Form
    {
        // Верхняя панель
        private TextBox txtDomain;

        // Вкладка Пользователь
        private TreeView tvUsers;
        private Label lblUsersDn;
        private Button btnRefreshUsersTree;
        private TextBox txtGivenName;
        private TextBox txtSurname;
        private TextBox txtSam;
        private Button btnSuggestSam;
        private TextBox txtUpnSuffix;
        private TextBox txtPassword;
        private Button btnGenPassword;
        private Button btnCopyPassword;
        private CheckBox chkShowPassword;
        private CheckBox chkMustChange;
        private Button btnCreateUser;
        private Button btnLoadUsersHere;

        // Вкладка Компьютер
        private TreeView tvComputers;
        private Label lblComputersDn;
        private Button btnRefreshComputersTree;
        private TextBox txtCompName;
        private Button btnCreateComputer;
        private Button btnLoadComputersHere;

        // Иконки
        private ImageList _icons;

        // Лог (без мигания)
        private RichTextBox txtLog;

        // Статус-бар
        private StatusStrip status;
        private ToolStripStatusLabel statusVersion;

        // Lazy-load
        private CancellationTokenSource _treeCts;
        private const string DummyNodeText = "…";

        // WM_SETREDRAW для безмигающего логирования
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        private const int WM_SETREDRAW = 0x000B;

        public MainForm()
        {
            // DPI + double buffer (без WS_EX_COMPOSITED)
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96, 96);
            SetStyle(ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint, true);
            UpdateStyles();

            Text = "AD Manager — пользователи и компьютеры";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 1200;
            Height = 800;

            ApplyTheme();

            // === Зелёная шапка ===
            var top = new Panel { Dock = DockStyle.Top, Height = 64, BackColor = Color.FromArgb(28, 164, 97) };
            Controls.Add(top);

            var caption = new Label
            {
                Text = "AD Manager",
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 16f, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(16, 16)
            };
            top.Controls.Add(caption);

            // Правая часть: домен + кнопка
            var domWrap = new Panel { Dock = DockStyle.Right, Width = 520, Padding = new Padding(0, 14, 16, 14) };
            top.Controls.Add(domWrap);

            var lblDom = new Label
            {
                Text = "Домен (FQDN):",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f),
                AutoSize = true,
                Dock = DockStyle.Top
            };
            domWrap.Controls.Add(lblDom);

            var domRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(0, 4, 0, 0) };
            domRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            domRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 48));
            domWrap.Controls.Add(domRow);

            txtDomain = new TextBox { Dock = DockStyle.Fill, PlaceholderText = "corp.contoso.com (необязательно)" };
            domRow.Controls.Add(txtDomain, 0, 0);

            var btnApplyDomain = new Button
            {
                Text = "✓",
                Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(22, 132, 79)
            };
            btnApplyDomain.FlatAppearance.BorderSize = 0;
            btnApplyDomain.Click += (s, e) =>
            {
                BuildOuTree(tvUsers, lblUsersDn, includeUsersCn: true, includeComputersCn: false);
                BuildOuTree(tvComputers, lblComputersDn, includeUsersCn: false, includeComputersCn: true);
            };
            domRow.Controls.Add(btnApplyDomain, 1, 0);

            // === Тулбар ===
            var toolbar = CreateToolbar();
            toolbar.Dock = DockStyle.Top;
            Controls.Add(toolbar);
            toolbar.BringToFront();
            top.BringToFront();

            // === Статус-бар ===
            status = new StatusStrip();
            statusVersion = new ToolStripStatusLabel($"Версия: {AppVersion.FullVersion}");
            status.Items.Add(statusVersion);
            Controls.Add(status);

            // === Центральная область: левый сайдбар + контент ===
            var body = new SplitContainer
            {
                Dock = DockStyle.Fill,
                SplitterDistance = 330,
                BorderStyle = BorderStyle.None,
                SplitterWidth = 2,
                BackColor = Color.FromArgb(34, 40, 49)
            };
            body.Panel1.BackColor = Color.FromArgb(34, 40, 49);
            body.Panel2.BackColor = Color.White;
            Controls.Add(body);
            body.BringToFront();

            // ----- Левый сайдбар (плоские тёмные вкладки) -----
            var sidebar = new FlatTabControl
            {
                Dock = DockStyle.Fill,
                DarkMode = true,
                ItemSize = new Size(160, 28)
            };
            body.Panel1.Padding = new Padding(8);
            body.Panel1.Controls.Add(sidebar);

            var tabLeftUsers = new TabPage("OU (пользователи)") { BackColor = Color.FromArgb(34, 40, 49) };
            var tabLeftComps = new TabPage("OU (компьютеры)") { BackColor = Color.FromArgb(34, 40, 49) };
            sidebar.TabPages.Add(tabLeftUsers);
            sidebar.TabPages.Add(tabLeftComps);

            // Левая вкладка — пользователи
            var leftUsersWrap = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(4) };
            leftUsersWrap.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            leftUsersWrap.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            leftUsersWrap.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tabLeftUsers.Controls.Add(leftUsersWrap);

            btnRefreshUsersTree = MakeGreenBtn("Обновить OU");
            btnRefreshUsersTree.Click += (s, e) => BuildOuTree(tvUsers, lblUsersDn, includeUsersCn: true, includeComputersCn: false);
            leftUsersWrap.Controls.Add(btnRefreshUsersTree, 0, 0);

            tvUsers = new TreeView
            {
                Dock = DockStyle.Fill,
                HideSelection = false,
                BorderStyle = BorderStyle.None,
                BackColor = Color.FromArgb(43, 49, 58),
                ForeColor = Color.Gainsboro
            };
            tvUsers.BeforeExpand += Tv_BeforeExpand_LoadChildren;
            tvUsers.AfterSelect += (s, e) =>
            {
                if (e.Node?.Tag is OuNodeTag tagOu) lblUsersDn.Text = "Выбрано: " + tagOu.DistinguishedName;
                else if (e.Node?.Tag is AccountNodeTag tagAcc) lblUsersDn.Text = $"Выбрано: {tagAcc.Kind}: {tagAcc.DistinguishedName}";
                else lblUsersDn.Text = "Выбрано: —";
            };
            // смена иконки открытой/закрытой папки
            tvUsers.AfterExpand += Tree_AfterExpandCollapse;
            tvUsers.AfterCollapse += Tree_AfterExpandCollapse;

            leftUsersWrap.Controls.Add(tvUsers, 0, 1);

            lblUsersDn = new Label { Text = "Выбрано: —", ForeColor = Color.WhiteSmoke, Dock = DockStyle.Fill, Padding = new Padding(4, 6, 4, 6) };
            leftUsersWrap.Controls.Add(lblUsersDn, 0, 2);

            // Левая вкладка — компьютеры
            var leftCompsWrap = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(4) };
            leftCompsWrap.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            leftCompsWrap.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            leftCompsWrap.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tabLeftComps.Controls.Add(leftCompsWrap);

            btnRefreshComputersTree = MakeGreenBtn("Обновить OU");
            btnRefreshComputersTree.Click += (s, e) => BuildOuTree(tvComputers, lblComputersDn, includeUsersCn: false, includeComputersCn: true);
            leftCompsWrap.Controls.Add(btnRefreshComputersTree, 0, 0);

            tvComputers = new TreeView
            {
                Dock = DockStyle.Fill,
                HideSelection = false,
                BorderStyle = BorderStyle.None,
                BackColor = Color.FromArgb(43, 49, 58),
                ForeColor = Color.Gainsboro
            };
            tvComputers.BeforeExpand += Tv_BeforeExpand_LoadChildren;
            tvComputers.AfterSelect += (s, e) =>
            {
                if (e.Node?.Tag is OuNodeTag tagOu) lblComputersDn.Text = "Выбрано: " + tagOu.DistinguishedName;
                else if (e.Node?.Tag is AccountNodeTag tagAcc) lblComputersDn.Text = $"Выбрано: {tagAcc.Kind}: {tagAcc.DistinguishedName}";
                else lblComputersDn.Text = "Выбрано: —";
            };
            tvComputers.AfterExpand += Tree_AfterExpandCollapse;
            tvComputers.AfterCollapse += Tree_AfterExpandCollapse;

            leftCompsWrap.Controls.Add(tvComputers, 0, 1);

            lblComputersDn = new Label { Text = "Выбрано: —", ForeColor = Color.WhiteSmoke, Dock = DockStyle.Fill, Padding = new Padding(4, 6, 4, 6) };
            leftCompsWrap.Controls.Add(lblComputersDn, 0, 2);

            // ----- Правая панель: плоские светлые вкладки -----
            var tabs = new FlatTabControl
            {
                Dock = DockStyle.Fill,
                DarkMode = false,
                ItemSize = new Size(120, 28),
                BackColor = Color.White
            };
            body.Panel2.Padding = new Padding(12);
            body.Panel2.Controls.Add(tabs);

            var tabUser = new TabPage("Пользователь") { BackColor = Color.White };
            var tabComp = new TabPage("Компьютер") { BackColor = Color.White };
            tabs.TabPages.Add(tabUser);
            tabs.TabPages.Add(tabComp);

            // Карточка: пользователь
            var cardUser = MakeCard("Создание пользователя");
            tabUser.Controls.Add(cardUser);

            var gridUser = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 2, AutoSize = true, Padding = new Padding(6) };
            gridUser.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            gridUser.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            cardUser.Controls.Add(gridUser);

            gridUser.Controls.Add(new Label { Text = "Имя:", AutoSize = true }, 0, 0);
            txtGivenName = new TextBox { Dock = DockStyle.Fill };
            gridUser.Controls.Add(txtGivenName, 1, 0);

            gridUser.Controls.Add(new Label { Text = "Фамилия:", AutoSize = true }, 0, 1);
            txtSurname = new TextBox { Dock = DockStyle.Fill };
            gridUser.Controls.Add(txtSurname, 1, 1);

            gridUser.Controls.Add(new Label { Text = "Логин (sAMAccountName):", AutoSize = true }, 0, 2);
            var samRow = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill, WrapContents = false };
            txtSam = new TextBox { Width = 220, PlaceholderText = "например, i.ivanov" };
            btnSuggestSam = MakeGreyBtn("Предложить", 120);
            btnSuggestSam.Click += (s, e) => AutoFillSam(true);
            samRow.Controls.AddRange(new Control[] { txtSam, btnSuggestSam });
            gridUser.Controls.Add(samRow, 1, 2);

            gridUser.Controls.Add(new Label { Text = "UPN-суффикс:", AutoSize = true }, 0, 3);
            txtUpnSuffix = new TextBox { Dock = DockStyle.Fill };
            gridUser.Controls.Add(txtUpnSuffix, 1, 3);

            gridUser.Controls.Add(new Label { Text = "Начальный пароль:", AutoSize = true }, 0, 4);
            var pwdRow = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill, WrapContents = false };
            txtPassword = new TextBox { Width = 220, UseSystemPasswordChar = true };
            btnGenPassword = MakeGreyBtn("Сгенерировать", 130);
            btnCopyPassword = MakeGreyBtn("Копировать", 110);
            chkShowPassword = new CheckBox { Text = "Показать", AutoSize = true };
            btnGenPassword.Click += (s, e) =>
            {
                var pwd = AdUtils.GenerateSecurePassword(16);
                txtPassword.Text = pwd;
                LogOk("Сгенерирован безопасный пароль (16 знаков).");
            };
            btnCopyPassword.Click += (s, e) =>
            {
                var p = txtPassword.Text ?? string.Empty;
                if (string.IsNullOrEmpty(p)) { LogErr("Пароль пуст — нечего копировать."); return; }
                Clipboard.SetText(p);
                LogOk("Пароль скопирован в буфер обмена.");
            };
            chkShowPassword.CheckedChanged += (s, e) => txtPassword.UseSystemPasswordChar = !chkShowPassword.Checked;
            pwdRow.Controls.AddRange(new Control[] { txtPassword, btnGenPassword, btnCopyPassword, chkShowPassword });
            gridUser.Controls.Add(pwdRow, 1, 4);

            chkMustChange = new CheckBox { Text = "Сменить пароль при первом входе", AutoSize = true, Dock = DockStyle.Top, Padding = new Padding(6) };
            cardUser.Controls.Add(chkMustChange);

            btnCreateUser = MakeGreenBtn("Создать пользователя");
            btnCreateUser.Click += BtnCreateUser_Click;
            cardUser.Controls.Add(btnCreateUser);

            btnLoadUsersHere = MakeGreyBtn("Показать объекты в выбранном узле", 260);
            btnLoadUsersHere.Click += async (s, e) =>
            {
                if (tvUsers.SelectedNode?.Tag is OuNodeTag)
                {
                    try { Cursor = Cursors.WaitCursor; await LoadAccountsForNodeAsync(tvUsers.SelectedNode, includeUsers: true, includeComputers: false, CancellationToken.None); }
                    catch (Exception ex) { LogErr(ex.Message); }
                    finally { Cursor = Cursors.Default; }
                }
            };
            cardUser.Controls.Add(btnLoadUsersHere);

            // Карточка: компьютер
            var cardComp = MakeCard("Создание компьютера");
            tabComp.Controls.Add(cardComp);

            var gridComp = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 2, AutoSize = true, Padding = new Padding(6) };
            gridComp.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            gridComp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            cardComp.Controls.Add(gridComp);

            gridComp.Controls.Add(new Label { Text = "Имя компьютера:", AutoSize = true }, 0, 0);
            txtCompName = new TextBox { Dock = DockStyle.Fill, PlaceholderText = "например, PC-001" };
            gridComp.Controls.Add(txtCompName, 1, 0);

            btnCreateComputer = MakeGreenBtn("Создать компьютер");
            btnCreateComputer.Click += BtnCreateComputer_Click;
            cardComp.Controls.Add(btnCreateComputer);

            btnLoadComputersHere = MakeGreyBtn("Показать объекты в выбранном узле", 260);
            btnLoadComputersHere.Click += async (s, e) =>
            {
                if (tvComputers.SelectedNode?.Tag is OuNodeTag)
                {
                    try { Cursor = Cursors.WaitCursor; await LoadAccountsForNodeAsync(tvComputers.SelectedNode, includeUsers: false, includeComputers: true, CancellationToken.None); }
                    catch (Exception ex) { LogErr(ex.Message); }
                    finally { Cursor = Cursors.Default; }
                }
            };
            cardComp.Controls.Add(btnLoadComputersHere);

            // === ЛОГ (RichTextBox) ===
            txtLog = new RichTextBox
            {
                Dock = DockStyle.Bottom,
                Height = 180,
                ReadOnly = true,
                DetectUrls = false,
                WordWrap = false,
                ScrollBars = RichTextBoxScrollBars.Both,
                BorderStyle = BorderStyle.None,
                BackColor = Color.White
            };
            Controls.Add(txtLog);

            // Иконки
            InitTreeIcons();
            tvUsers.ImageList = _icons;
            tvComputers.ImageList = _icons;

            // Значения по умолчанию
            txtUpnSuffix.Text = AdUtils.GetDefaultUpnSuffix();

            // Стартовая загрузка
            Shown += (s, e) =>
            {
                BuildOuTree(tvUsers, lblUsersDn, includeUsersCn: true, includeComputersCn: false);
                BuildOuTree(tvComputers, lblComputersDn, includeUsersCn: false, includeComputersCn: true);
            };

            // Автогенерация логина
            txtGivenName.TextChanged += (s, e) => AutoFillSam();
            txtSurname.TextChanged += (s, e) => AutoFillSam();
        }

        // ------------------------- UI helpers -------------------------
        private void ApplyTheme()
        {
            Font = new Font("Segoe UI", 10f);
            ToolStripManager.Renderer = new GreenRenderer();
            BackColor = Color.White;
        }

        private ToolStrip CreateToolbar()
        {
            var ts = new ToolStrip
            {
                GripStyle = ToolStripGripStyle.Hidden,
                ImageScalingSize = new Size(20, 20),
                BackColor = Color.FromArgb(246, 248, 250)
            };
            ts.Items.Add(new ToolStripButton("Обновить все", null, (s, e) =>
            {
                BuildOuTree(tvUsers, lblUsersDn, includeUsersCn: true, includeComputersCn: false);
                BuildOuTree(tvComputers, lblComputersDn, includeUsersCn: false, includeComputersCn: true);
            })
            { DisplayStyle = ToolStripItemDisplayStyle.Text });

            ts.Items.Add(new ToolStripSeparator());
            ts.Items.Add(new ToolStripLabel("Функции:"));
            ts.Items.Add(new ToolStripButton("Создать пользователя") { DisplayStyle = ToolStripItemDisplayStyle.Text });
            ts.Items.Add(new ToolStripButton("Создать компьютер") { DisplayStyle = ToolStripItemDisplayStyle.Text });
            return ts;
        }

        private Button MakeGreenBtn(string text)
        {
            var b = new Button
            {
                Text = text,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(46, 196, 109),
                ForeColor = Color.White,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(10, 6, 10, 6),
                Margin = new Padding(6),
                UseCompatibleTextRendering = true,
                MinimumSize = new Size(180, 0)
            };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }

        private Button MakeGreyBtn(string text, int minWidth = 120)
        {
            var b = new Button
            {
                Text = text,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(235, 240, 245),
                ForeColor = Color.Black,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(10, 6, 10, 6),
                Margin = new Padding(6, 0, 0, 0),
                UseCompatibleTextRendering = true,
                MinimumSize = new Size(minWidth, 0)
            };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }

        private FlowLayoutPanel MakeCard(string title)
        {
            var card = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(12),
                Margin = new Padding(0, 0, 0, 12),
                BackColor = Color.White
            };
            card.Paint += (s, e) =>
            {
                var r = new Rectangle(0, 0, card.Width - 1, card.Height - 1);
                using var p = new Pen(Color.FromArgb(230, 234, 238));
                e.Graphics.DrawRectangle(p, r);
            };

            var header = new Label
            {
                Text = title,
                Font = new Font("Segoe UI Semibold", 12f),
                AutoSize = true,
                ForeColor = Color.FromArgb(33, 37, 41),
                Margin = new Padding(0, 0, 0, 8)
            };
            card.Controls.Add(header);
            return card;
        }

        private class GreenRenderer : ToolStripProfessionalRenderer
        {
            protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) { }
            protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e)
            {
                var rect = new Rectangle(Point.Empty, e.Item.Size);
                if (e.Item.Selected || (e.Item as ToolStripButton)?.Checked == true)
                    using (var b = new SolidBrush(Color.FromArgb(228, 245, 236)))
                        e.Graphics.FillRectangle(b, rect);
            }
        }

        // ------------------------- ИКОНКИ -------------------------
        private void InitTreeIcons()
        {
            _icons = new ImageList
            {
                ColorDepth = ColorDepth.Depth32Bit,
                ImageSize = new Size(16, 16)
            };

            // пытаемся загрузить из файлов проекта
            Bitmap TryLoad(string fileName, Bitmap fallback)
            {
                try
                {
                    var p1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", fileName);
                    var p2 = Path.Combine(Environment.CurrentDirectory, "Resources", fileName);
                    var path = File.Exists(p1) ? p1 : (File.Exists(p2) ? p2 : null);
                    if (path != null)
                    {
                        using var tmp = (Bitmap)Image.FromFile(path);
                        return new Bitmap(tmp); // копия, чтобы файл не держался открытым
                    }
                }
                catch { }
                return fallback;
            }

            var icoDomain = TryLoad("Domain16.png", SystemIcons.Shield.ToBitmap());
            var icoFolder = TryLoad("FolderClosed16.png", SystemIcons.Application.ToBitmap());
            var icoFolderOpen = TryLoad("FolderOpen16.png", SystemIcons.Application.ToBitmap());
            var icoUser = TryLoad("User16.png", SystemIcons.Information.ToBitmap());
            var icoComputer = TryLoad("Computer16.png", SystemIcons.WinLogo.ToBitmap());

            _icons.Images.Add("domain", icoDomain);
            _icons.Images.Add("ou", icoFolder);
            _icons.Images.Add("ou_open", icoFolderOpen);
            _icons.Images.Add("container", icoFolder);
            _icons.Images.Add("container_open", icoFolderOpen);
            _icons.Images.Add("user", icoUser);
            _icons.Images.Add("computer", icoComputer);
        }

        // динамическая смена папки (открыта/закрыта)
        private void Tree_AfterExpandCollapse(object sender, TreeViewEventArgs e)
        {
            if (e.Node?.Tag is OuNodeTag tag)
            {
                switch (tag.Kind)
                {
                    case NodeKind.OrganizationalUnit:
                        e.Node.ImageKey = (e.Action == TreeViewAction.Expand) ? "ou_open" : "ou";
                        e.Node.SelectedImageKey = e.Node.ImageKey;
                        break;
                    case NodeKind.Container:
                        e.Node.ImageKey = (e.Action == TreeViewAction.Expand) ? "container_open" : "container";
                        e.Node.SelectedImageKey = e.Node.ImageKey;
                        break;
                }
            }
        }

        // ------------------------- ЛОГИРОВАНИЕ (без мерцания) -------------------------
        private void AppendLogLine(string text, bool ok)
        {
            if (txtLog.IsDisposed) return;

            try
            {
                SendMessage(txtLog.Handle, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero); // freeze

                var timestamp = $"[{DateTime.Now:T}] ";
                txtLog.SelectionStart = txtLog.TextLength;
                txtLog.SelectionLength = 0;

                txtLog.SelectionColor = Color.DimGray;
                txtLog.AppendText(timestamp);

                txtLog.SelectionColor = ok ? Color.ForestGreen : Color.Firebrick;
                txtLog.AppendText(ok ? "✅ " : "❌ ");

                txtLog.SelectionColor = Color.Black;
                txtLog.AppendText(text + Environment.NewLine);
            }
            finally
            {
                SendMessage(txtLog.Handle, WM_SETREDRAW, new IntPtr(1), IntPtr.Zero); // unfreeze
                txtLog.Invalidate();
                txtLog.SelectionStart = txtLog.TextLength;
                txtLog.ScrollToCaret();
            }
        }
        private void LogOk(string message) => AppendLogLine(message, ok: true);
        private void LogErr(string message) => AppendLogLine("Ошибка: " + message, ok: false);

        // ------------------------- AD-логика (как у тебя) -------------------------
        private void BtnCreateUser_Click(object sender, EventArgs e)
        {
            AppendLogLine("Создание пользователя...", true);
            try
            {
                var domain = AdUtils.SafeTrim(txtDomain.Text);

                var node = tvUsers.SelectedNode;
                if (node?.Tag is not OuNodeTag ouTagUser)
                    throw new ArgumentException("Выберите OU/контейнер для пользователя в дереве слева.");
                var ouDn = ouTagUser.DistinguishedName;

                var given = AdUtils.SafeTrim(txtGivenName.Text);
                var sur = AdUtils.SafeTrim(txtSurname.Text);
                var samInput = AdUtils.SafeTrim(txtSam.Text);
                var upnSuffix = AdUtils.SafeTrim(txtUpnSuffix.Text);
                var password = txtPassword.Text ?? string.Empty;
                var mustChange = chkMustChange.Checked;

                if (string.IsNullOrWhiteSpace(given)) throw new ArgumentException("Не указано имя.");
                if (string.IsNullOrWhiteSpace(sur)) throw new ArgumentException("Не указана фамилия.");
                if (string.IsNullOrWhiteSpace(upnSuffix)) throw new ArgumentException("Не указан UPN-суффикс (например, contoso.com).");
                if (string.IsNullOrWhiteSpace(password)) throw new ArgumentException("Не указан начальный пароль.");

                var baseSam = !string.IsNullOrWhiteSpace(samInput) ? samInput : AdUtils.BuildBestGuessSam(given, sur);

                using (var ctx = AdUtils.BuildPrincipalContext(domain, ouDn))
                {
                    if (ctx == null)
                        throw new InvalidOperationException("Не удалось создать контекст домена. Проверьте домен и OU DN.");

                    var sam = AdUtils.GenerateUniqueSam(ctx, baseSam, given, sur, maxLen: 20, log: msg => LogOk(msg));
                    var upn = $"{sam}@{upnSuffix}";

                    if (AdUtils.FindUser(ctx, sam) != null)
                        throw new InvalidOperationException($"Неожиданно: пользователь '{sam}' уже существует.");

                    var user = new UserPrincipal(ctx)
                    {
                        SamAccountName = sam,
                        GivenName = given,
                        Surname = sur,
                        DisplayName = $"{given} {sur}",
                        UserPrincipalName = upn,
                        Enabled = false
                    };

                    user.Save();
                    user.SetPassword(password);
                    if (mustChange) user.ExpirePasswordNow();
                    user.Enabled = true;
                    user.Save();

                    LogOk($"Пользователь '{sam}' создан в '{ouDn}'. UPN: {upn}");
                }
            }
            catch (PasswordException ex) { LogErr("Пароль не соответствует политике домена: " + ex.Message); }
            catch (PrincipalExistsException ex) { LogErr("Объект уже существует: " + ex.Message); }
            catch (Exception ex) { LogErr(ex.Message); }
        }

        private void BtnCreateComputer_Click(object sender, EventArgs e)
        {
            AppendLogLine("Создание компьютера...", true);
            try
            {
                var domain = AdUtils.SafeTrim(txtDomain.Text);
                var compName = AdUtils.SafeTrim(txtCompName.Text);
                if (string.IsNullOrWhiteSpace(compName)) throw new ArgumentException("Не указано имя компьютера.");

                var node = tvComputers.SelectedNode;
                if (node?.Tag is not OuNodeTag ouTagComp)
                    throw new ArgumentException("Выберите OU/контейнер для компьютера в дереве слева.");

                var containerDn = ouTagComp.DistinguishedName;
                var sam = compName.EndsWith("$") ? compName : compName + "$";

                using (var ctx = AdUtils.BuildPrincipalContext(domain, containerDn))
                {
                    if (ctx == null)
                        throw new InvalidOperationException("Не удалось создать контекст домена. Проверьте домен и OU DN.");

                    if (AdUtils.FindComputer(ctx, compName) != null)
                        throw new InvalidOperationException($"Компьютер '{compName}' уже существует.");

                    var computer = new ComputerPrincipal(ctx)
                    {
                        SamAccountName = sam,
                        Name = compName,
                        Enabled = true
                    };
                    computer.Save();

                    LogOk($"Компьютер '{compName}' создан в '{containerDn}'.");
                }
            }
            catch (PrincipalExistsException ex) { LogErr("Объект уже существует: " + ex.Message); }
            catch (Exception ex) { LogErr(ex.Message); }
        }

        private async void BuildOuTree(TreeView tv, Label lblSelected, bool includeUsersCn, bool includeComputersCn)
        {
            _treeCts?.Cancel();
            _treeCts = new CancellationTokenSource();
            var ct = _treeCts.Token;

            try
            {
                tv.BeginUpdate();
                tv.Nodes.Clear();
                lblSelected.Text = "Выбрано: —";

                var domainFqdn = AdUtils.SafeTrim(txtDomain.Text);
                var defaultNc = GetDefaultNamingContext(domainFqdn);
                if (string.IsNullOrWhiteSpace(defaultNc))
                    throw new InvalidOperationException("Не удалось определить defaultNamingContext.");

                var rootNode = new TreeNode(defaultNc)
                {
                    Tag = new OuNodeTag { Name = defaultNc, DistinguishedName = defaultNc, Kind = NodeKind.DomainRoot },
                    ImageKey = "domain",
                    SelectedImageKey = "domain"
                };
                rootNode.Nodes.Add(new TreeNode(DummyNodeText)); // стрелка
                tv.Nodes.Add(rootNode);
                tv.EndUpdate();

                await LoadChildrenOneLevelAsync(rootNode, includeUsersCn, includeComputersCn, ct);
                rootNode.Expand();
                LogOk("Структура OU загружена (ленивая подгрузка).");
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                tv.EndUpdate();
                LogErr("Не удалось загрузить дерево OU: " + ex.Message);
            }
        }

        private async void Tv_BeforeExpand_LoadChildren(object sender, TreeViewCancelEventArgs e)
        {
            if (e.Node.Nodes.Count == 1 && e.Node.Nodes[0].Text == DummyNodeText)
            {
                try
                {
                    Cursor = Cursors.AppStarting;
                    var tv = (TreeView)sender;
                    bool includeUsers = (tv == tvUsers);
                    bool includeComputers = (tv == tvComputers);

                    _treeCts?.Cancel();
                    _treeCts = new CancellationTokenSource();
                    var ct = _treeCts.Token;

                    await LoadChildrenOneLevelAsync(e.Node, includeUsers, includeComputers, ct);
                }
                catch (OperationCanceledException) { }
                catch (Exception ex) { LogErr("Ошибка подгрузки узла: " + ex.Message); }
                finally { Cursor = Cursors.Default; }
            }
        }

        private async Task LoadChildrenOneLevelAsync(TreeNode parentNode, bool includeUsersCn, bool includeComputersCn, CancellationToken ct)
        {
            if (parentNode.Nodes.Count == 1 && parentNode.Nodes[0].Text == DummyNodeText)
                parentNode.Nodes.Clear();

            if (parentNode.Tag is not OuNodeTag tag || string.IsNullOrWhiteSpace(tag.DistinguishedName))
                return;

            var domainFqdn = AdUtils.SafeTrim(txtDomain.Text);

            var children = await Task.Run(() =>
                AdUtils.FetchChildContainers(domainFqdn, tag.DistinguishedName), ct);

            parentNode.TreeView.BeginUpdate();
            try
            {
                foreach (var child in children)
                {
                    TreeNode node;
                    if (child.Kind == NodeKind.OrganizationalUnit)
                    {
                        node = new TreeNode($"{child.Name} (OU)")
                        {
                            Tag = child,
                            ImageKey = "ou",
                            SelectedImageKey = "ou_open"
                        };
                    }
                    else
                    {
                        node = new TreeNode($"{child.Name} (CN)")
                        {
                            Tag = child,
                            ImageKey = "container",
                            SelectedImageKey = "container_open"
                        };
                    }

                    node.Nodes.Add(new TreeNode(DummyNodeText)); // возможны дети
                    parentNode.Nodes.Add(node);
                }

                if (tag.Kind == NodeKind.DomainRoot)
                {
                    if (includeUsersCn)
                    {
                        var cnUsers = AdUtils.TryGetNamedContainer(domainFqdn, tag.DistinguishedName, "CN=Users");
                        if (cnUsers != null)
                        {
                            var n = new TreeNode($"{cnUsers.Name} (CN)")
                            {
                                Tag = cnUsers,
                                ImageKey = "container",
                                SelectedImageKey = "container_open"
                            };
                            n.Nodes.Add(new TreeNode(DummyNodeText));
                            parentNode.Nodes.Add(n);
                        }
                    }
                    if (includeComputersCn)
                    {
                        var cnComputers = AdUtils.TryGetNamedContainer(domainFqdn, tag.DistinguishedName, "CN=Computers");
                        if (cnComputers != null)
                        {
                            var n = new TreeNode($"{cnComputers.Name} (CN)")
                            {
                                Tag = cnComputers,
                                ImageKey = "container",
                                SelectedImageKey = "container_open"
                            };
                            n.Nodes.Add(new TreeNode(DummyNodeText));
                            parentNode.Nodes.Add(n);
                        }
                    }
                }
            }
            finally
            {
                parentNode.TreeView.EndUpdate();
            }
        }

        private async Task LoadAccountsForNodeAsync(TreeNode node, bool includeUsers, bool includeComputers, CancellationToken ct)
        {
            if (node?.Tag is not OuNodeTag tag) return;

            var domainFqdn = AdUtils.SafeTrim(txtDomain.Text);

            // Удаляем ранее подгруженных "детей-аккаунтов"
            var toRemove = new List<TreeNode>();
            foreach (TreeNode ch in node.Nodes)
                if (ch.Tag is AccountNodeTag) toRemove.Add(ch);
            foreach (var r in toRemove) node.Nodes.Remove(r);

            var accs = await Task.Run(() =>
                AdUtils.FetchAccounts(domainFqdn, tag.DistinguishedName, includeUsers, includeComputers), ct);

            node.TreeView.BeginUpdate();
            try
            {
                foreach (var a in accs)
                {
                    var imageKey = a.Kind == NodeKind.User ? "user" : "computer";
                    var accNode = new TreeNode(a.Name)
                    {
                        Tag = a,
                        ImageKey = imageKey,
                        SelectedImageKey = imageKey
                    };
                    node.Nodes.Add(accNode);
                }
            }
            finally
            {
                node.TreeView.EndUpdate();
            }
        }

        private string GetDefaultNamingContext(string domainFqdn)
        {
            try
            {
                DirectoryEntry rootDse = string.IsNullOrWhiteSpace(domainFqdn)
                    ? new DirectoryEntry("LDAP://RootDSE")
                    : new DirectoryEntry($"LDAP://{domainFqdn}/RootDSE");

                return rootDse.Properties["defaultNamingContext"]?.Value as string;
            }
            catch { return null; }
        }

        private void AutoFillSam(bool force = false)
        {
            var given = AdUtils.SafeTrim(txtGivenName.Text);
            var sur = AdUtils.SafeTrim(txtSurname.Text);
            if (string.IsNullOrWhiteSpace(given) || string.IsNullOrWhiteSpace(sur)) return;
            if (!force && !string.IsNullOrWhiteSpace(AdUtils.SafeTrim(txtSam.Text))) return;
            txtSam.Text = AdUtils.BuildBestGuessSam(given, sur);
        }
    }
}
