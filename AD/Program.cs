using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Drawing;

namespace AD
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        // Верхняя панель
        private TextBox txtDomain; // corp.contoso.com (опционально)

        // Вкладка Пользователь
        private TreeView tvUsers;           // Иерархический выбор OU (пользователи)
        private Label lblUsersDn;           // Показ DN выбранного контейнера/объекта
        private Button btnRefreshUsersTree;
        private TextBox txtGivenName;
        private TextBox txtSurname;
        private TextBox txtSam;
        private Button btnSuggestSam;       // Предложить логин
        private TextBox txtUpnSuffix;
        private TextBox txtPassword;
        private Button btnGenPassword;      // Сгенерировать пароль
        private Button btnCopyPassword;     // Копировать пароль
        private CheckBox chkShowPassword;   // Показать/скрыть пароль
        private CheckBox chkMustChange;
        private Button btnCreateUser;

        // Вкладка Компьютер
        private TreeView tvComputers;       // Иерархический выбор OU (компьютеры)
        private Label lblComputersDn;       // Показ DN выбранного контейнера/объекта
        private Button btnRefreshComputersTree;
        private TextBox txtCompName;
        private Button btnCreateComputer;

        // Иконки
        private ImageList _icons;

        // Лог
        private TextBox txtLog;

        public MainForm()
        {
            Text = "AD Manager — пользователи и компьютеры (иерархический выбор OU)";
            Width = 1100;
            Height = 750;
            StartPosition = FormStartPosition.CenterScreen;

            // === Корневой лэйаут ===
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // общая панель
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 65));// вкладки
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 35));// лог
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

            // === Вкладки ===
            var tabs = new TabControl { Dock = DockStyle.Fill };
            root.Controls.Add(tabs, 0, 1);

            var tabUser = new TabPage("Пользователь");
            var tabComp = new TabPage("Компьютер");
            tabs.TabPages.Add(tabUser);
            tabs.TabPages.Add(tabComp);

            // === Вкладка Пользователь ===
            var splitUser = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 420
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

            tvUsers = new TreeView
            {
                Dock = DockStyle.Fill,
                HideSelection = false
            };
            tvUsers.AfterSelect += (s, e) =>
            {
                if (e.Node?.Tag is OuNodeTag tagOu)
                    lblUsersDn.Text = "Выбрано: " + tagOu.DistinguishedName;
                else if (e.Node?.Tag is AccountNodeTag tagAcc)
                    lblUsersDn.Text = $"Выбрано: {tagAcc.Kind}: {tagAcc.DistinguishedName}";
                else
                    lblUsersDn.Text = "Выбрано: —";
            };
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

            // Пароль + кнопки "Сгенерировать", "Копировать", + чекбокс "Показать"
            pnlUserForm.Controls.Add(new Label { Text = "Начальный пароль:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 4);
            var pnlPwd = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            txtPassword = new TextBox { Width = 220, UseSystemPasswordChar = true };
            btnGenPassword = new Button { Text = "Сгенерировать", AutoSize = true };
            btnCopyPassword = new Button { Text = "Копировать", AutoSize = true };
            chkShowPassword = new CheckBox { Text = "Показать пароль", AutoSize = true };

            btnGenPassword.Click += (s, e) =>
            {
                var pwd = GenerateSecurePassword(16);
                txtPassword.Text = pwd;
                LogOk("Сгенерирован безопасный пароль (16 знаков).");
            };

            btnCopyPassword.Click += (s, e) =>
            {
                try
                {
                    var p = txtPassword.Text ?? string.Empty;
                    if (string.IsNullOrEmpty(p))
                    {
                        LogErr("Пароль пуст — нечего копировать.");
                        return;
                    }
                    Clipboard.SetText(p);
                    LogOk("Пароль скопирован в буфер обмена.");
                }
                catch (Exception ex)
                {
                    LogErr("Не удалось скопировать пароль: " + ex.Message);
                }
            };

            chkShowPassword.CheckedChanged += (s, e) =>
            {
                txtPassword.UseSystemPasswordChar = !chkShowPassword.Checked;
            };

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

            // Автогенерация login при вводе ФИО
            txtGivenName.TextChanged += (s, e) => AutoFillSam();
            txtSurname.TextChanged += (s, e) => AutoFillSam();

            splitUser.Panel2.Controls.Add(pnlUserForm);

            // === Вкладка Компьютер ===
            var splitComp = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 420
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

            tvComputers = new TreeView
            {
                Dock = DockStyle.Fill,
                HideSelection = false
            };
            tvComputers.AfterSelect += (s, e) =>
            {
                if (e.Node?.Tag is OuNodeTag tagOu)
                    lblComputersDn.Text = "Выбрано: " + tagOu.DistinguishedName;
                else if (e.Node?.Tag is AccountNodeTag tagAcc)
                    lblComputersDn.Text = $"Выбрано: {tagAcc.Kind}: {tagAcc.DistinguishedName}";
                else
                    lblComputersDn.Text = "Выбрано: —";
            };
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

            splitComp.Panel2.Controls.Add(pnlCompForm);

            // === ЛОГ ===
            txtLog = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, WordWrap = false };
            root.Controls.Add(txtLog, 0, 2);

            // Значения по умолчанию
            txtUpnSuffix.Text = GetDefaultUpnSuffix();

            // Иконки и привязка к деревьям
            InitTreeIcons();
            tvUsers.ImageList = _icons;
            tvComputers.ImageList = _icons;

            // При первом показе — построить деревья OU
            Shown += (s, e) =>
            {
                BuildOuTree(tvUsers, lblUsersDn, includeUsersCn: true, includeComputersCn: false);
                BuildOuTree(tvComputers, lblComputersDn, includeUsersCn: false, includeComputersCn: true);
            };
        }

        private void InitTreeIcons()
        {
            _icons = new ImageList { ColorDepth = ColorDepth.Depth32Bit, ImageSize = new Size(16, 16) };
            // Заглушки: при желании замените на свои ресурсы
            _icons.Images.Add("domain", SystemIcons.Shield.ToBitmap());
            _icons.Images.Add("ou", SystemIcons.Application.ToBitmap());
            _icons.Images.Add("container", SystemIcons.Asterisk.ToBitmap());
            _icons.Images.Add("user", SystemIcons.Information.ToBitmap());
            _icons.Images.Add("computer", SystemIcons.WinLogo.ToBitmap());
            _icons.Images.Add("default", SystemIcons.Application.ToBitmap());
        }

        private string GetDefaultUpnSuffix()
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(Environment.UserDomainName))
                {
                    return Environment.UserDomainName.Contains(".")
                        ? Environment.UserDomainName
                        : Environment.UserDomainName + ".local";
                }
            }
            catch { }
            return "contoso.local";
        }

        // ===== Создание пользователя =====
        private void BtnCreateUser_Click(object sender, EventArgs e)
        {
            txtLog.AppendText($"[{DateTime.Now:T}] Создание пользователя...\r\n");
            try
            {
                var domain = SafeTrim(txtDomain.Text);

                var node = tvUsers.SelectedNode;
                if (node?.Tag is not OuNodeTag ouTagUser)
                    throw new ArgumentException("Выберите OU/контейнер для пользователя в дереве слева.");
                var ouDn = ouTagUser.DistinguishedName;

                var given = SafeTrim(txtGivenName.Text);
                var sur = SafeTrim(txtSurname.Text);
                var samInput = SafeTrim(txtSam.Text);
                var upnSuffix = SafeTrim(txtUpnSuffix.Text);
                var password = txtPassword.Text ?? string.Empty;
                var mustChange = chkMustChange.Checked;

                if (string.IsNullOrWhiteSpace(given)) throw new ArgumentException("Не указано имя.");
                if (string.IsNullOrWhiteSpace(sur)) throw new ArgumentException("Не указана фамилия.");
                if (string.IsNullOrWhiteSpace(upnSuffix)) throw new ArgumentException("Не указан UPN-суффикс (например, contoso.com).");
                if (string.IsNullOrWhiteSpace(password)) throw new ArgumentException("Не указан начальный пароль.");

                // Предлагаем логин, если поле пустое
                var baseSam = !string.IsNullOrWhiteSpace(samInput) ? samInput : BuildBestGuessSam(given, sur);

                using (var ctx = BuildPrincipalContext(domain, ouDn))
                {
                    if (ctx == null)
                        throw new InvalidOperationException("Не удалось создать контекст домена. Проверьте домен и OU DN.");

                    // Получаем уникальный sAMAccountName (если занято — добавляем варианты/цифры)
                    var sam = GenerateUniqueSam(ctx, baseSam, given, sur, maxLen: 20, log: msg => LogOk(msg));

                    var upn = $"{sam}@{upnSuffix}";

                    if (FindUser(ctx, sam) != null)
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

                    // Устанавливаем пароль отдельно, чтобы ловить политику сложности домена
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
                var domain = SafeTrim(txtDomain.Text);
                var compName = SafeTrim(txtCompName.Text);
                if (string.IsNullOrWhiteSpace(compName)) throw new ArgumentException("Не указано имя компьютера.");

                var node = tvComputers.SelectedNode;
                if (node?.Tag is not OuNodeTag ouTagComp)
                    throw new ArgumentException("Выберите OU/контейнер для компьютера в дереве слева.");

                var containerDn = ouTagComp.DistinguishedName;
                var sam = compName.EndsWith("$") ? compName : compName + "$";

                using (var ctx = BuildPrincipalContext(domain, containerDn))
                {
                    if (ctx == null)
                        throw new InvalidOperationException("Не удалось создать контекст домена. Проверьте домен и OU DN.");

                    if (FindComputer(ctx, compName) != null)
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

        // ===== Построение дерева OU + аккаунты =====
        private void BuildOuTree(TreeView tv, Label lblSelected, bool includeUsersCn, bool includeComputersCn)
        {
            try
            {
                tv.BeginUpdate();
                tv.Nodes.Clear();
                lblSelected.Text = "Выбрано: —";

                var domainFqdn = SafeTrim(txtDomain.Text);

                // Получаем defaultNamingContext
                var defaultNc = GetDefaultNamingContext(domainFqdn);
                if (string.IsNullOrWhiteSpace(defaultNc))
                    throw new InvalidOperationException("Не удалось определить defaultNamingContext. Проверьте подключение к домену.");

                // Корень домена как корневой узел дерева
                var rootNode = new TreeNode(defaultNc)
                {
                    Tag = new OuNodeTag { Name = defaultNc, DistinguishedName = defaultNc, Kind = NodeKind.DomainRoot },
                    ImageKey = "domain",
                    SelectedImageKey = "domain"
                };
                tv.Nodes.Add(rootNode);

                // Рекурсивно добавляем OU
                using (var rootDe = new DirectoryEntry(BuildLdapPath(domainFqdn, defaultNc)))
                {
                    // Спец-контейнеры
                    TreeNode usersCnNode = null;
                    TreeNode computersCnNode = null;

                    if (includeUsersCn) usersCnNode = TryAddContainerNode(rootDe, rootNode, "CN=Users");
                    if (includeComputersCn) computersCnNode = TryAddContainerNode(rootDe, rootNode, "CN=Computers");

                    if (usersCnNode != null) AddAccountsToNode(rootDe, usersCnNode, includeUsers: true, includeComputers: false);
                    if (computersCnNode != null) AddAccountsToNode(rootDe, computersCnNode, includeUsers: false, includeComputers: true);

                    // OU
                    foreach (DirectoryEntry ouChild in rootDe.Children)
                    {
                        if (IsObjectClass(ouChild, "organizationalUnit"))
                        {
                            var ouNode = CreateOuTreeNode(ouChild);
                            rootNode.Nodes.Add(ouNode);

                            // Добавляем аккаунты в этом OU
                            AddAccountsToNode(ouChild, ouNode, includeUsersCn, includeComputersCn);

                            // Идём глубже
                            PopulateOuChildrenRecursively(ouChild, ouNode, includeUsersCn, includeComputersCn);
                        }
                    }
                }

                rootNode.Expand();
                tv.EndUpdate();
                LogOk("Дерево OU загружено.");
            }
            catch (Exception ex)
            {
                tv.EndUpdate();
                LogErr("Не удалось загрузить дерево OU: " + ex.Message);
            }
        }

        private void PopulateOuChildrenRecursively(DirectoryEntry parentDe, TreeNode parentNode, bool includeUsersCn, bool includeComputersCn)
        {
            foreach (DirectoryEntry child in parentDe.Children)
            {
                try
                {
                    if (IsObjectClass(child, "organizationalUnit"))
                    {
                        var ouNode = CreateOuTreeNode(child);
                        parentNode.Nodes.Add(ouNode);

                        // Загружаем пользователей/компьютеры уровня этого OU
                        AddAccountsToNode(child, ouNode, includeUsersCn, includeComputersCn);

                        PopulateOuChildrenRecursively(child, ouNode, includeUsersCn, includeComputersCn);
                    }
                    else if (IsObjectClass(child, "container"))
                    {
                        // иногда админы делают контейнеры (container) вместо OU
                        var dn = SafeProp(child, "distinguishedName");
                        var name = SafeProp(child, "name");
                        var node = new TreeNode($"{name} (CN)")
                        {
                            Tag = new OuNodeTag { Name = name, DistinguishedName = dn, Kind = NodeKind.Container },
                            ImageKey = "container",
                            SelectedImageKey = "container"
                        };
                        parentNode.Nodes.Add(node);

                        // подгружаем аккаунты в этом контейнере
                        AddAccountsToNode(child, node, includeUsersCn, includeComputersCn);
                    }
                    else if (includeUsersCn && NameEquals(child, "CN=Users"))
                    {
                        var cnNode = TryAddContainerNode(parentDe, parentNode, "CN=Users");
                        if (cnNode != null) AddAccountsToNode(child, cnNode, includeUsers: true, includeComputers: false);
                    }
                    else if (includeComputersCn && NameEquals(child, "CN=Computers"))
                    {
                        var cnNode = TryAddContainerNode(parentDe, parentNode, "CN=Computers");
                        if (cnNode != null) AddAccountsToNode(child, cnNode, includeUsers: false, includeComputers: true);
                    }
                }
                catch
                {
                    // пропускаем проблемные ветки
                }
            }
        }

        private TreeNode CreateOuTreeNode(DirectoryEntry de)
        {
            var dn = SafeProp(de, "distinguishedName");
            var name = SafeProp(de, "name");
            var node = new TreeNode($"{name} (OU)")
            {
                Tag = new OuNodeTag { Name = name, DistinguishedName = dn, Kind = NodeKind.OrganizationalUnit },
                ImageKey = "ou",
                SelectedImageKey = "ou"
            };
            return node;
        }

        private TreeNode TryAddContainerNode(DirectoryEntry parent, TreeNode parentNode, string cnName)
        {
            try
            {
                using var ds = new DirectorySearcher(parent)
                {
                    Filter = $"(&(objectClass=container)(name={EscapeLdapValue(cnName.Split('=')[1])}))",
                    SearchScope = SearchScope.OneLevel,
                    PageSize = 50
                };
                ds.PropertiesToLoad.Add("name");
                ds.PropertiesToLoad.Add("distinguishedName");
                var res = ds.FindOne();
                if (res != null)
                {
                    var dn = res.Properties["distinguishedName"]?[0]?.ToString();
                    var name = res.Properties["name"]?[0]?.ToString();
                    if (!string.IsNullOrWhiteSpace(dn))
                    {
                        var node = new TreeNode($"{name} (CN)")
                        {
                            Tag = new OuNodeTag { Name = name ?? dn, DistinguishedName = dn, Kind = NodeKind.Container },
                            ImageKey = "container",
                            SelectedImageKey = "container"
                        };
                        parentNode.Nodes.Add(node);
                        return node;
                    }
                }
            }
            catch { /* необязательно */ }
            return null;
        }

        private void AddAccountsToNode(DirectoryEntry containerDe, TreeNode parentNode, bool includeUsers, bool includeComputers)
        {
            try
            {
                using var ds = new DirectorySearcher(containerDe)
                {
                    SearchScope = SearchScope.OneLevel,
                    PageSize = 200
                };

                if (includeUsers)
                {
                    ds.Filter = "(&(|(objectCategory=person))(objectClass=user))";
                    ds.PropertiesToLoad.Clear();
                    ds.PropertiesToLoad.Add("name");
                    ds.PropertiesToLoad.Add("displayName");
                    ds.PropertiesToLoad.Add("sAMAccountName");
                    ds.PropertiesToLoad.Add("distinguishedName");

                    foreach (SearchResult r in ds.FindAll())
                    {
                        var dn = r.Properties["distinguishedName"]?.Count > 0 ? r.Properties["distinguishedName"][0]?.ToString() : null;
                        var sam = r.Properties["sAMAccountName"]?.Count > 0 ? r.Properties["sAMAccountName"][0]?.ToString() : null;
                        var disp = r.Properties["displayName"]?.Count > 0 ? r.Properties["displayName"][0]?.ToString() : null;
                        var nm = r.Properties["name"]?.Count > 0 ? r.Properties["name"][0]?.ToString() : null;

                        var shown = !string.IsNullOrWhiteSpace(disp) ? disp : nm ?? sam ?? "(user)";
                        var userNode = new TreeNode(shown)
                        {
                            Tag = new AccountNodeTag
                            {
                                Name = shown,
                                SamAccountName = sam,
                                DistinguishedName = dn,
                                Kind = NodeKind.User
                            },
                            ImageKey = "user",
                            SelectedImageKey = "user"
                        };
                        parentNode.Nodes.Add(userNode);
                    }
                }

                if (includeComputers)
                {
                    ds.Filter = "(objectClass=computer)";
                    ds.PropertiesToLoad.Clear();
                    ds.PropertiesToLoad.Add("name");
                    ds.PropertiesToLoad.Add("sAMAccountName");
                    ds.PropertiesToLoad.Add("distinguishedName");

                    foreach (SearchResult r in ds.FindAll())
                    {
                        var dn = r.Properties["distinguishedName"]?.Count > 0 ? r.Properties["distinguishedName"][0]?.ToString() : null;
                        var sam = r.Properties["sAMAccountName"]?.Count > 0 ? r.Properties["sAMAccountName"][0]?.ToString() : null;
                        var nm = r.Properties["name"]?.Count > 0 ? r.Properties["name"][0]?.ToString() : null;

                        var shown = (nm ?? sam ?? "(computer)")?.TrimEnd('$');

                        var compNode = new TreeNode(shown)
                        {
                            Tag = new AccountNodeTag
                            {
                                Name = shown,
                                SamAccountName = sam,
                                DistinguishedName = dn,
                                Kind = NodeKind.Computer
                            },
                            ImageKey = "computer",
                            SelectedImageKey = "computer"
                        };
                        parentNode.Nodes.Add(compNode);
                    }
                }
            }
            catch (Exception ex)
            {
                LogErr("Не удалось загрузить объекты в узел: " + ex.Message);
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

        // ===== AD контекст/поиск =====
        private PrincipalContext BuildPrincipalContext(string domainFqdn, string containerDn)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(domainFqdn))
                {
                    if (!string.IsNullOrWhiteSpace(containerDn))
                        return new PrincipalContext(ContextType.Domain, domainFqdn, containerDn);
                    else
                        return new PrincipalContext(ContextType.Domain, domainFqdn);
                }
                else
                {
                    if (!string.IsNullOrWhiteSpace(containerDn))
                        return new PrincipalContext(ContextType.Domain, null, containerDn);
                    else
                        return new PrincipalContext(ContextType.Domain);
                }
            }
            catch (Exception ex)
            {
                LogErr("Ошибка создания контекста домена: " + ex.Message);
                return null;
            }
        }

        private UserPrincipal FindUser(PrincipalContext ctx, string sam)
        {
            try { return UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, sam); } catch { return null; }
        }

        private ComputerPrincipal FindComputer(PrincipalContext ctx, string compName)
        {
            try { return ComputerPrincipal.FindByIdentity(ctx, IdentityType.Name, compName); } catch { return null; }
        }

        // ===== Автогенерация логина (ФИО → sAMAccountName) =====

        private void AutoFillSam(bool force = false)
        {
            var given = SafeTrim(txtGivenName.Text);
            var sur = SafeTrim(txtSurname.Text);
            if (string.IsNullOrWhiteSpace(given) || string.IsNullOrWhiteSpace(sur))
                return;

            // если уже руками ввели логин — не перезатираем, кроме случая force
            if (!force && !string.IsNullOrWhiteSpace(SafeTrim(txtSam.Text))) return;

            var sam = BuildBestGuessSam(given, sur);
            txtSam.Text = sam;
        }

        private string BuildBestGuessSam(string given, string surname)
        {
            var g = TransliterateRuToEn(given);
            var s = TransliterateRuToEn(surname);

            g = SlugifyAscii(g);
            s = SlugifyAscii(s);

            // Базовый вариант: i.surname
            var initial = g.Length > 0 ? g[0].ToString() : "";
            var cand = (initial.Length > 0 ? $"{initial}.{s}" : s);

            cand = TrimSamToMax(cand).Trim('.');
            if (string.IsNullOrWhiteSpace(cand)) cand = "user";

            return cand.ToLowerInvariant();
        }

        private string GenerateUniqueSam(PrincipalContext ctx, string baseSam, string given, string surname, int maxLen, Action<string> log)
        {
            string Normalize(string x) => TrimSamToMax(SlugifyAscii(x).Trim('.')).ToLowerInvariant();

            var g = Normalize(TransliterateRuToEn(given));
            var s = Normalize(TransliterateRuToEn(surname));
            var i = g.Length > 0 ? g[0].ToString() : "";

            var baseNormalized = Normalize(baseSam);

            var candidates = new List<string>();
            void AddIf(string x)
            {
                x = TrimSamToMax(x).Trim('.');
                if (!string.IsNullOrWhiteSpace(x)) candidates.Add(x.ToLowerInvariant());
            }

            AddIf(baseNormalized);
            if (!string.IsNullOrEmpty(i) && !string.IsNullOrEmpty(s)) AddIf($"{i}.{s}");
            if (!string.IsNullOrEmpty(g) && !string.IsNullOrEmpty(s)) AddIf($"{g}.{s}");
            if (!string.IsNullOrEmpty(s)) AddIf($"{s}");
            if (!string.IsNullOrEmpty(s) && !string.IsNullOrEmpty(i)) AddIf($"{s}.{i}");
            if (!string.IsNullOrEmpty(i) && !string.IsNullOrEmpty(s)) AddIf($"{i}{s}");
            if (!string.IsNullOrEmpty(s) && !string.IsNullOrEmpty(i)) AddIf($"{s}{i}");

            foreach (var cand in candidates.Distinct())
            {
                if (FindUser(ctx, cand) == null) return cand;
            }

            var baseCore = candidates.FirstOrDefault() ?? baseNormalized;
            baseCore = TrimSamToMax(baseCore);

            for (int n = 1; n <= 9999; n++)
            {
                var suffix = n.ToString();
                var trimmed = baseCore;
                if (trimmed.Length + suffix.Length > maxLen)
                    trimmed = trimmed.Substring(0, Math.Max(0, maxLen - suffix.Length));

                var attempt = (trimmed + suffix).Trim('.').ToLowerInvariant();
                if (attempt.Length == 0) continue;

                if (FindUser(ctx, attempt) == null)
                {
                    log?.Invoke($"Логин '{baseSam}' занят. Выбран свободный вариант: '{attempt}'.");
                    return attempt;
                }
            }

            throw new InvalidOperationException("Не удалось подобрать свободный sAMAccountName.");
        }

        // Транслитерация RU→EN (упрощённая)
        private static string TransliterateRuToEn(string src)
        {
            if (string.IsNullOrWhiteSpace(src)) return "";
            var map = new Dictionary<char, string>
            {
                ['а'] = "a",
                ['б'] = "b",
                ['в'] = "v",
                ['г'] = "g",
                ['д'] = "d",
                ['е'] = "e",
                ['ё'] = "e",
                ['ж'] = "zh",
                ['з'] = "z",
                ['и'] = "i",
                ['й'] = "y",
                ['к'] = "k",
                ['л'] = "l",
                ['м'] = "m",
                ['н'] = "n",
                ['о'] = "o",
                ['п'] = "p",
                ['р'] = "r",
                ['с'] = "s",
                ['т'] = "t",
                ['у'] = "u",
                ['ф'] = "f",
                ['х'] = "h",
                ['ц'] = "c",
                ['ч'] = "ch",
                ['ш'] = "sh",
                ['щ'] = "sch",
                ['ъ'] = "",
                ['ы'] = "y",
                ['ь'] = "",
                ['э'] = "e",
                ['ю'] = "yu",
                ['я'] = "ya",
                ['А'] = "a",
                ['Б'] = "b",
                ['В'] = "v",
                ['Г'] = "g",
                ['Д'] = "d",
                ['Е'] = "e",
                ['Ё'] = "e",
                ['Ж'] = "zh",
                ['З'] = "z",
                ['И'] = "i",
                ['Й'] = "y",
                ['К'] = "k",
                ['Л'] = "l",
                ['М'] = "m",
                ['Н'] = "n",
                ['О'] = "o",
                ['П'] = "p",
                ['Р'] = "r",
                ['С'] = "s",
                ['Т'] = "t",
                ['У'] = "u",
                ['Ф'] = "f",
                ['Х'] = "h",
                ['Ц'] = "c",
                ['Ч'] = "ch",
                ['Ш'] = "sh",
                ['Щ'] = "sch",
                ['Ъ'] = "",
                ['Ы'] = "y",
                ['Ь'] = "",
                ['Э'] = "e",
                ['Ю'] = "yu",
                ['Я'] = "ya",
            };

            var sb = new StringBuilder(src.Length * 2);
            foreach (var ch in src)
            {
                if (map.TryGetValue(ch, out var rep)) sb.Append(rep);
                else sb.Append(ch);
            }
            return sb.ToString();
        }

        // Оставляем только [a-z0-9.-], lower; заменяем пробелы/подчёркивания на точку
        private static string SlugifyAscii(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            s = s.ToLowerInvariant();
            s = s.Replace(' ', '.').Replace('_', '.');

            var sb = new StringBuilder(s.Length);
            foreach (var ch in s)
            {
                if ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9') || ch == '.' || ch == '-')
                    sb.Append(ch);
            }

            var res = Regex.Replace(sb.ToString(), @"\.{2,}", ".");
            res = res.Trim('.');
            return res;
        }

        private static string TrimSamToMax(string s, int max = 20)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return s.Length <= max ? s : s.Substring(0, max);
        }

        // ===== Генерация безопасного пароля =====
        private static string GenerateSecurePassword(int length = 16)
        {
            if (length < 12) length = 12; // разумный минимум

            const string U = "ABCDEFGHJKLMNPQRSTUVWXYZ"; // без I, O
            const string L = "abcdefghijkmnopqrstuvwxyz"; // без l
            const string D = "23456789";                  // без 0,1
            const string S = "!@#$%^&*()-_=+[]{};:,.?/"; // без кавычек/бэкслеша

            var sets = new[] { U, L, D, S };

            var chars = new List<char>();
            foreach (var set in sets)
                chars.Add(GetRandomChar(set));

            string all = string.Concat(sets);
            while (chars.Count < length)
            {
                var c = GetRandomChar(all);
                if (chars.Count > 0 && chars[^1] == c) continue;
                chars.Add(c);
            }

            Shuffle(chars);

            return new string(chars.ToArray());
        }

        private static char GetRandomChar(string allowed)
        {
            var idx = GetRandomInt(0, allowed.Length);
            return allowed[idx];
        }

        private static int GetRandomInt(int minInclusive, int maxExclusive)
        {
            if (minInclusive >= maxExclusive) return minInclusive;
            var diff = (long)maxExclusive - minInclusive;
            var uint32Buffer = new byte[4];
            using (var rng = RandomNumberGenerator.Create())
            {
                while (true)
                {
                    rng.GetBytes(uint32Buffer);
                    uint rand = BitConverter.ToUInt32(uint32Buffer, 0);
                    long max = (1 + (long)uint.MaxValue);
                    long remainder = max % diff;
                    if (rand < max - remainder)
                        return (int)(minInclusive + (rand % diff));
                }
            }
        }

        private static void Shuffle<T>(IList<T> list)
        {
            using var rng = RandomNumberGenerator.Create();
            for (int i = list.Count - 1; i > 0; i--)
            {
                int j = GetRandomInt(0, i + 1);
                (list[i], list[j]) = (list[j], list[i]);
            }
        }

        // ===== Вспомогательные =====
        private static string SafeTrim(string s) => string.IsNullOrWhiteSpace(s) ? string.Empty : s.Trim();
        private void LogOk(string message) => txtLog.AppendText($"[{DateTime.Now:T}] ✅ {message}\r\n");
        private void LogErr(string message) => txtLog.AppendText($"[{DateTime.Now:T}] ❌ Ошибка: {message}\r\n");

        private static string BuildLdapPath(string domainFqdn, string dn)
        {
            var prefix = string.IsNullOrWhiteSpace(domainFqdn) ? "LDAP://" : $"LDAP://{domainFqdn}/";
            return prefix + dn;
        }

        private static bool IsObjectClass(DirectoryEntry de, string className)
        {
            try
            {
                var oc = de.Properties["objectClass"];
                if (oc == null) return false;
                foreach (var v in oc) if (v?.ToString()?.Equals(className, StringComparison.OrdinalIgnoreCase) == true) return true;
                return false;
            }
            catch { return false; }
        }

        private static bool NameEquals(DirectoryEntry de, string expected)
        {
            try
            {
                var n = de.Properties["name"]?.Value?.ToString();
                return string.Equals("name=" + n, expected, StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(n, expected.split('=')[1], StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        private static string SafeProp(DirectoryEntry de, string prop)
        {
            try { return de.Properties[prop]?.Value?.ToString() ?? string.Empty; } catch { return string.Empty; }
        }

        private static string EscapeLdapValue(string value)
        {
            if (value == null) return string.Empty;
            var sb = new StringBuilder();
            foreach (var c in value)
            {
                switch (c)
                {
                    case '\\': sb.Append(@"\5c"); break;
                    case '*': sb.Append(@"\2a"); break;
                    case '(': sb.Append(@"\28"); break;
                    case ')': sb.Append(@"\29"); break;
                    case '\0': sb.Append(@"\00"); break;
                    case '/': sb.Append(@"\2f"); break;
                    default: sb.Append(c); break;
                }
            }
            return sb.ToString();
        }

        // Узлы дерева
        private class OuNodeTag
        {
            public string Name { get; set; }
            public string DistinguishedName { get; set; }
            public NodeKind Kind { get; set; }
        }

        private class AccountNodeTag
        {
            public string Name { get; set; }
            public string SamAccountName { get; set; }
            public string DistinguishedName { get; set; }
            public NodeKind Kind { get; set; } // User / Computer
        }

        private enum NodeKind
        {
            DomainRoot,
            OrganizationalUnit,
            Container,
            User,
            Computer
        }
    }
}
