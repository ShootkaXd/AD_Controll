using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AD
{
    public class MainForm : Form
    {
        // Верхняя панель
        private TextBox txtDomain; // corp.contoso.com (опционально)

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
        private Button btnLoadUsersHere;    // показать объекты в выбранном узле

        // Вкладка Компьютер
        private TreeView tvComputers;
        private Label lblComputersDn;
        private Button btnRefreshComputersTree;
        private TextBox txtCompName;
        private Button btnCreateComputer;
        private Button btnLoadComputersHere;// показать объекты в выбранном узле

        // Иконки
        private ImageList _icons;

        // Лог
        private TextBox txtLog;

        // Версия
        private StatusStrip status;
        private ToolStripStatusLabel statusVersion;

        // Lazy-load / async
        private CancellationTokenSource _treeCts;
        private const string DummyNodeText = "…";

        public MainForm()
        {
            Text = "AD Manager — пользователи и компьютеры";
            Width = 1150;
            Height = 780;
            StartPosition = FormStartPosition.CenterScreen;

            // === Главное меню ===
            var menu = new MenuStrip();
            var miFile = new ToolStripMenuItem("Файл");
            var miExit = new ToolStripMenuItem("Выход", null, (s, e) => Close());
            miFile.DropDownItems.Add(miExit);

            var miHelp = new ToolStripMenuItem("Справка");
            var miAbout = new ToolStripMenuItem("О программе…", null, (s, e) =>
            {
                MessageBox.Show($"AD Manager\n{AppVersion.FullVersion}\n© {DateTime.Now:yyyy}",
                    "О программе", MessageBoxButtons.OK, MessageBoxIcon.Information);
            });

            miHelp.DropDownItems.Add(miAbout);
            menu.Items.Add(miFile);
            menu.Items.Add(miHelp);
            MainMenuStrip = menu;
            Controls.Add(menu);

            // === Статус-бар с версией ===
            status = new StatusStrip();
            statusVersion = new ToolStripStatusLabel($"Версия: {AppVersion.FullVersion}");
            status.Items.Add(statusVersion);
            Controls.Add(status);

            // === Корневой лэйаут ===
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // общая панель
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));// вкладки + лог
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // статус-бар (уже добавлен)
            Controls.Add(root);

            // === Общая панель сверху ===
            var pnlCommon = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                AutoSize = true,
                Padding = new Padding(8),
            };
            pnlCommon.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pnlCommon.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            pnlCommon.Controls.Add(new Label { Text = "Домен (FQDN):", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            txtDomain = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, PlaceholderText = "corp.contoso.com (необязательно)" };
            pnlCommon.Controls.Add(txtDomain, 1, 0);

            root.Controls.Add(pnlCommon, 0, 0);

            // === Средняя зона: вкладки + лог ===
            var mid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
            };
            mid.RowStyles.Add(new RowStyle(SizeType.Percent, 64));
            mid.RowStyles.Add(new RowStyle(SizeType.Percent, 36));
            root.Controls.Add(mid, 0, 1);

            // === Вкладки ===
            var tabs = new TabControl { Dock = DockStyle.Fill };
            mid.Controls.Add(tabs, 0, 0);

            var tabUser = new TabPage("Пользователь");
            var tabComp = new TabPage("Компьютер");
            tabs.TabPages.Add(tabUser);
            tabs.TabPages.Add(tabComp);

            // === Вкладка Пользователь ===
            var splitUser = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 430
            };
            tabUser.Controls.Add(splitUser);

            // Левая панель — дерево OU (пользователи)
            var leftUser = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(8)
            };
            leftUser.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            leftUser.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            leftUser.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            btnRefreshUsersTree = new Button { Text = "Обновить OU", AutoSize = true };
            btnRefreshUsersTree.Click += (s, e) => BuildOuTree(tvUsers, lblUsersDn, includeUsersCn: true, includeComputersCn: false);
            leftUser.Controls.Add(btnRefreshUsersTree, 0, 0);

            tvUsers = new TreeView { Dock = DockStyle.Fill, HideSelection = false };
            tvUsers.AfterSelect += (s, e) =>
            {
                if (e.Node?.Tag is OuNodeTag tagOu)
                    lblUsersDn.Text = "Выбрано: " + tagOu.DistinguishedName;
                else if (e.Node?.Tag is AccountNodeTag tagAcc)
                    lblUsersDn.Text = $"Выбрано: {tagAcc.Kind}: {tagAcc.DistinguishedName}";
                else
                    lblUsersDn.Text = "Выбрано: —";
            };
            tvUsers.BeforeExpand += Tv_BeforeExpand_LoadChildren;
            leftUser.Controls.Add(tvUsers, 0, 1);

            lblUsersDn = new Label { Text = "Выбрано: —", AutoSize = true, Padding = new Padding(0, 6, 0, 0) };
            leftUser.Controls.Add(lblUsersDn, 0, 2);

            splitUser.Panel1.Controls.Add(leftUser);

            // Правая панель — форма создания пользователя
            var pnlUserForm = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                Padding = new Padding(8)
            };
            pnlUserForm.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pnlUserForm.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            pnlUserForm.Controls.Add(new Label { Text = "Имя:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            txtGivenName = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            pnlUserForm.Controls.Add(txtGivenName, 1, 0);

            pnlUserForm.Controls.Add(new Label { Text = "Фамилия:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
            txtSurname = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            pnlUserForm.Controls.Add(txtSurname, 1, 1);

            // Логин + кнопка "Предложить логин"
            pnlUserForm.Controls.Add(new Label { Text = "Логин (sAMAccountName):", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
            var pnlSam = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            txtSam = new TextBox { Width = 220, PlaceholderText = "например, i.ivanov" };
            btnSuggestSam = new Button { Text = "Предложить логин", AutoSize = true };
            btnSuggestSam.Click += (s, e) => AutoFillSam(force: true);
            pnlSam.Controls.Add(txtSam);
            pnlSam.Controls.Add(btnSuggestSam);
            pnlUserForm.Controls.Add(pnlSam, 1, 2);

            pnlUserForm.Controls.Add(new Label { Text = "UPN-суффикс:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
            txtUpnSuffix = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, PlaceholderText = "например, contoso.com" };
            pnlUserForm.Controls.Add(txtUpnSuffix, 1, 3);

            // Пароль + кнопки
            pnlUserForm.Controls.Add(new Label { Text = "Начальный пароль:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 4);
            var pnlPwd = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            txtPassword = new TextBox { Width = 220, UseSystemPasswordChar = true };
            btnGenPassword = new Button { Text = "Сгенерировать", AutoSize = true };
            btnCopyPassword = new Button { Text = "Копировать", AutoSize = true };
            chkShowPassword = new CheckBox { Text = "Показать пароль", AutoSize = true };

            btnGenPassword.Click += (s, e) =>
            {
                var pwd = AdUtils.GenerateSecurePassword(16);
                txtPassword.Text = pwd;
                LogOk("Сгенерирован безопасный пароль (16 знаков).");
            };
            btnCopyPassword.Click += (s, e) =>
            {
                try
                {
                    var p = txtPassword.Text ?? string.Empty;
                    if (string.IsNullOrEmpty(p)) { LogErr("Пароль пуст — нечего копировать."); return; }
                    Clipboard.SetText(p);
                    LogOk("Пароль скопирован в буфер обмена.");
                }
                catch (Exception ex) { LogErr("Не удалось скопировать пароль: " + ex.Message); }
            };
            chkShowPassword.CheckedChanged += (s, e) => { txtPassword.UseSystemPasswordChar = !chkShowPassword.Checked; };

            pnlPwd.Controls.Add(txtPassword);
            pnlPwd.Controls.Add(btnGenPassword);
            pnlPwd.Controls.Add(btnCopyPassword);
            pnlPwd.Controls.Add(chkShowPassword);
            pnlUserForm.Controls.Add(pnlPwd, 1, 4);

            chkMustChange = new CheckBox { Text = "Сменить пароль при первом входе", AutoSize = true };
            pnlUserForm.Controls.Add(chkMustChange, 1, 5);

            btnCreateUser = new Button { Text = "Создать пользователя", AutoSize = true };
            btnCreateUser.Click += BtnCreateUser_Click;
            pnlUserForm.Controls.Add(btnCreateUser, 1, 6);

            // Кнопка «Показать объекты в выбранном узле»
            btnLoadUsersHere = new Button { Text = "Показать объекты в выбранном узле", AutoSize = true };
            btnLoadUsersHere.Click += async (s, e) =>
            {
                if (tvUsers.SelectedNode?.Tag is OuNodeTag)
                {
                    try
                    {
                        Cursor = Cursors.WaitCursor;
                        await LoadAccountsForNodeAsync(tvUsers.SelectedNode, includeUsers: true, includeComputers: false, CancellationToken.None);
                    }
                    catch (Exception ex) { LogErr(ex.Message); }
                    finally { Cursor = Cursors.Default; }
                }
            };
            pnlUserForm.Controls.Add(btnLoadUsersHere, 1, 7);

            // Автогенерация login при вводе ФИО
            txtGivenName.TextChanged += (s, e) => AutoFillSam();
            txtSurname.TextChanged += (s, e) => AutoFillSam();

            splitUser.Panel2.Controls.Add(pnlUserForm);

            // === Вкладка Компьютер ===
            var splitComp = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 430
            };
            tabComp.Controls.Add(splitComp);

            // Левая панель — дерево OU (компьютеры)
            var leftComp = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(8)
            };
            leftComp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            leftComp.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            leftComp.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            btnRefreshComputersTree = new Button { Text = "Обновить OU", AutoSize = true };
            btnRefreshComputersTree.Click += (s, e) => BuildOuTree(tvComputers, lblComputersDn, includeUsersCn: false, includeComputersCn: true);
            leftComp.Controls.Add(btnRefreshComputersTree, 0, 0);

            tvComputers = new TreeView { Dock = DockStyle.Fill, HideSelection = false };
            tvComputers.AfterSelect += (s, e) =>
            {
                if (e.Node?.Tag is OuNodeTag tagOu)
                    lblComputersDn.Text = "Выбрано: " + tagOu.DistinguishedName;
                else if (e.Node?.Tag is AccountNodeTag tagAcc)
                    lblComputersDn.Text = $"Выбрано: {tagAcc.Kind}: {tagAcc.DistinguishedName}";
                else
                    lblComputersDn.Text = "Выбрано: —";
            };
            tvComputers.BeforeExpand += Tv_BeforeExpand_LoadChildren;
            leftComp.Controls.Add(tvComputers, 0, 1);

            lblComputersDn = new Label { Text = "Выбрано: —", AutoSize = true, Padding = new Padding(0, 6, 0, 0) };
            leftComp.Controls.Add(lblComputersDn, 0, 2);

            splitComp.Panel1.Controls.Add(leftComp);

            // Правая панель — форма создания компьютера
            var pnlCompForm = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                Padding = new Padding(8)
            };
            pnlCompForm.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pnlCompForm.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            pnlCompForm.Controls.Add(new Label { Text = "Имя компьютера:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            txtCompName = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, PlaceholderText = "например, PC-001" };
            pnlCompForm.Controls.Add(txtCompName, 1, 0);

            btnCreateComputer = new Button { Text = "Создать компьютер", AutoSize = true };
            btnCreateComputer.Click += BtnCreateComputer_Click;
            pnlCompForm.Controls.Add(btnCreateComputer, 1, 1);

            // Кнопка «Показать объекты в выбранном узле»
            btnLoadComputersHere = new Button { Text = "Показать объекты в выбранном узле", AutoSize = true };
            btnLoadComputersHere.Click += async (s, e) =>
            {
                if (tvComputers.SelectedNode?.Tag is OuNodeTag)
                {
                    try
                    {
                        Cursor = Cursors.WaitCursor;
                        await LoadAccountsForNodeAsync(tvComputers.SelectedNode, includeUsers: false, includeComputers: true, CancellationToken.None);
                    }
                    catch (Exception ex) { LogErr(ex.Message); }
                    finally { Cursor = Cursors.Default; }
                }
            };
            pnlCompForm.Controls.Add(btnLoadComputersHere, 1, 2);

            splitComp.Panel2.Controls.Add(pnlCompForm);

            // === ЛОГ ===
            txtLog = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, WordWrap = false };
            mid.Controls.Add(txtLog, 0, 1);

            // Значения по умолчанию
            txtUpnSuffix.Text = AdUtils.GetDefaultUpnSuffix();

            // Иконки и привязка к деревьям
            InitTreeIcons();
            tvUsers.ImageList = _icons;
            tvComputers.ImageList = _icons;

            // При первом показе — построить деревья OU (лениво)
            Shown += (s, e) =>
            {
                BuildOuTree(tvUsers, lblUsersDn, includeUsersCn: true, includeComputersCn: false);
                BuildOuTree(tvComputers, lblComputersDn, includeUsersCn: false, includeComputersCn: true);
            };
        }

        private void InitTreeIcons()
        {
            _icons = new ImageList { ColorDepth = ColorDepth.Depth32Bit, ImageSize = new Size(16, 16) };
            _icons.Images.Add("domain", SystemIcons.Shield.ToBitmap());
            _icons.Images.Add("ou", SystemIcons.Application.ToBitmap());
            _icons.Images.Add("container", SystemIcons.Asterisk.ToBitmap());
            _icons.Images.Add("user", SystemIcons.Information.ToBitmap());
            _icons.Images.Add("computer", SystemIcons.WinLogo.ToBitmap());
            _icons.Images.Add("default", SystemIcons.Application.ToBitmap());
        }

        // ===== Создание пользователя =====
        private void BtnCreateUser_Click(object sender, EventArgs e)
        {
            txtLog.AppendText($"[{DateTime.Now:T}] Создание пользователя...\r\n");
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

        // ===== Создание компьютера =====
        private void BtnCreateComputer_Click(object sender, EventArgs e)
        {
            txtLog.AppendText($"[{DateTime.Now:T}] Создание компьютера...\r\n");
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

        // ===== Ленивая загрузка дерева =====
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
                tv.Nodes.Add(rootNode);

                // Добавляем «заглушку», чтобы появилась стрелка раскрытия
                rootNode.Nodes.Add(new TreeNode(DummyNodeText));
                tv.EndUpdate();

                // Подгружаем только первый уровень
                await LoadChildrenOneLevelAsync(rootNode, includeUsersCn, includeComputersCn, ct);
                rootNode.Expand();
                LogOk("Структура OU загружена (ленивая подгрузка).");
            }
            catch (OperationCanceledException) { /* ignore */ }
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
                catch (OperationCanceledException) { /* ignore */ }
                catch (Exception ex)
                {
                    LogErr("Ошибка подгрузки узла: " + ex.Message);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
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
                            SelectedImageKey = "ou"
                        };
                    }
                    else
                    {
                        node = new TreeNode($"{child.Name} (CN)")
                        {
                            Tag = child,
                            ImageKey = "container",
                            SelectedImageKey = "container"
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
                                SelectedImageKey = "container"
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
                                SelectedImageKey = "container"
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
                DirectoryEntry rootDse;
                if (!string.IsNullOrWhiteSpace(domainFqdn))
                    rootDse = new DirectoryEntry($"LDAP://{domainFqdn}/RootDSE");
                else
                    rootDse = new DirectoryEntry("LDAP://RootDSE");

                return rootDse.Properties["defaultNamingContext"]?.Value as string;
            }
            catch
            {
                return null;
            }
        }

        // ===== Автогенерация логина (ФИО → sAMAccountName) =====
        private void AutoFillSam(bool force = false)
        {
            var given = AdUtils.SafeTrim(txtGivenName.Text);
            var sur = AdUtils.SafeTrim(txtSurname.Text);
            if (string.IsNullOrWhiteSpace(given) || string.IsNullOrWhiteSpace(sur))
                return;

            if (!force && !string.IsNullOrWhiteSpace(AdUtils.SafeTrim(txtSam.Text))) return;

            var sam = AdUtils.BuildBestGuessSam(given, sur);
            txtSam.Text = sam;
        }

        // ===== Логирование =====
        private void LogOk(string message) => txtLog.AppendText($"[{DateTime.Now:T}] ✅ {message}\r\n");
        private void LogErr(string message) => txtLog.AppendText($"[{DateTime.Now:T}] ❌ Ошибка: {message}\r\n");
    }
}
