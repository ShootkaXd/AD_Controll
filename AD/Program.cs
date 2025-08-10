using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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
        // Верхняя панель (общие параметры)
        private TextBox txtDomain;       // corp.contoso.com (опционально)

        // Вкладка Пользователь — НОВОЕ: список OU
        private ComboBox cboOuUsers;     // список OU/контейнеров для пользователей
        private Button btnRefreshUsersOUs;
        private TextBox txtGivenName;
        private TextBox txtSurname;
        private TextBox txtSam;
        private TextBox txtUpnSuffix;    // contoso.com
        private TextBox txtPassword;
        private CheckBox chkMustChange;
        private Button btnCreateUser;

        // Вкладка Компьютер — список OU
        private ComboBox cboOuComputers; // список OU/контейнеров для компьютеров
        private Button btnRefreshComputersOUs;
        private TextBox txtCompName;
        private Button btnCreateComputer;

        // Лог
        private TextBox txtLog;

        public MainForm()
        {
            Text = "AD Manager — пользователи и компьютеры";
            Width = 1000;
            Height = 720;
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
            var pnlUser = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                AutoSize = true,
                Padding = new Padding(8),
            };
            pnlUser.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pnlUser.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            pnlUser.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pnlUser.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            // НОВОЕ: выбор OU для пользователей
            pnlUser.Controls.Add(new Label { Text = "OU для пользователей:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            cboOuUsers = new ComboBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, DropDownStyle = ComboBoxStyle.DropDownList };
            pnlUser.Controls.Add(cboOuUsers, 1, 0);
            pnlUser.SetColumnSpan(cboOuUsers, 2);
            btnRefreshUsersOUs = new Button { Text = "Обновить список OU", AutoSize = true };
            btnRefreshUsersOUs.Click += (s, e) => LoadOuListForUsers();
            pnlUser.Controls.Add(btnRefreshUsersOUs, 3, 0);

            pnlUser.Controls.Add(new Label { Text = "Имя:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
            txtGivenName = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            pnlUser.Controls.Add(txtGivenName, 1, 1);

            pnlUser.Controls.Add(new Label { Text = "Фамилия:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 1);
            txtSurname = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            pnlUser.Controls.Add(txtSurname, 3, 1);

            pnlUser.Controls.Add(new Label { Text = "Логин (sAMAccountName):", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
            txtSam = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, PlaceholderText = "например, j.doe" };
            pnlUser.Controls.Add(txtSam, 1, 2);

            pnlUser.Controls.Add(new Label { Text = "UPN‑суффикс:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 2);
            txtUpnSuffix = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, PlaceholderText = "например, contoso.com" };
            pnlUser.Controls.Add(txtUpnSuffix, 3, 2);

            pnlUser.Controls.Add(new Label { Text = "Начальный пароль:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
            txtPassword = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, UseSystemPasswordChar = true };
            pnlUser.Controls.Add(txtPassword, 1, 3);

            chkMustChange = new CheckBox { Text = "Сменить пароль при первом входе", AutoSize = true };
            pnlUser.Controls.Add(chkMustChange, 2, 3);
            pnlUser.SetColumnSpan(chkMustChange, 2);

            btnCreateUser = new Button { Text = "Создать пользователя", AutoSize = true };
            btnCreateUser.Click += BtnCreateUser_Click;
            pnlUser.Controls.Add(btnCreateUser, 0, 4);
            pnlUser.SetColumnSpan(btnCreateUser, 4);

            tabUser.Controls.Add(pnlUser);

            // === Вкладка Компьютер ===
            var pnlComp = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                AutoSize = true,
                Padding = new Padding(8),
            };
            pnlComp.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pnlComp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            pnlComp.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            pnlComp.Controls.Add(new Label { Text = "OU для компьютеров:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            cboOuComputers = new ComboBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, DropDownStyle = ComboBoxStyle.DropDownList };
            pnlComp.Controls.Add(cboOuComputers, 1, 0);
            btnRefreshComputersOUs = new Button { Text = "Обновить список OU", AutoSize = true };
            btnRefreshComputersOUs.Click += (s, e) => LoadOuListForComputers();
            pnlComp.Controls.Add(btnRefreshComputersOUs, 2, 0);

            pnlComp.Controls.Add(new Label { Text = "Имя компьютера:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
            txtCompName = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, PlaceholderText = "например, PC-001" };
            pnlComp.Controls.Add(txtCompName, 1, 1);
            pnlComp.SetColumnSpan(txtCompName, 2);

            btnCreateComputer = new Button { Text = "Создать компьютер", AutoSize = true };
            btnCreateComputer.Click += BtnCreateComputer_Click;
            pnlComp.Controls.Add(btnCreateComputer, 0, 2);
            pnlComp.SetColumnSpan(btnCreateComputer, 3);

            tabComp.Controls.Add(pnlComp);

            // === ЛОГ ===
            txtLog = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true, WordWrap = false };
            root.Controls.Add(txtLog, 0, 2);

            // Значения по умолчанию
            txtUpnSuffix.Text = GetDefaultUpnSuffix();

            // При первом показе — загрузить OU для обеих вкладок
            Shown += (s, e) => { LoadOuListForUsers(); LoadOuListForComputers(); };
        }

        private string GetDefaultUpnSuffix()
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(Environment.UserDomainName))
                {
                    return Environment.UserDomainName.Contains(".") ? Environment.UserDomainName : Environment.UserDomainName + ".local";
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
                if (cboOuUsers.SelectedItem is not OuItem ouUser)
                    throw new ArgumentException("Выберите OU для пользователя.");
                var ouDn = ouUser.DistinguishedName;

                var given = SafeTrim(txtGivenName.Text);
                var sur = SafeTrim(txtSurname.Text);
                var sam = SafeTrim(txtSam.Text);
                var upnSuffix = SafeTrim(txtUpnSuffix.Text);
                var password = txtPassword.Text ?? string.Empty;
                var mustChange = chkMustChange.Checked;

                if (string.IsNullOrWhiteSpace(sam)) throw new ArgumentException("Не указан логин (sAMAccountName).");
                if (string.IsNullOrWhiteSpace(given)) throw new ArgumentException("Не указано имя.");
                if (string.IsNullOrWhiteSpace(sur)) throw new ArgumentException("Не указана фамилия.");
                if (string.IsNullOrWhiteSpace(password)) throw new ArgumentException("Не указан начальный пароль.");
                if (string.IsNullOrWhiteSpace(upnSuffix)) throw new ArgumentException("Не указан UPN‑суффикс (например, contoso.com).");

                var upn = $"{sam}@{upnSuffix}";

                using (var ctx = BuildPrincipalContext(domain, ouDn))
                {
                    if (ctx == null)
                        throw new InvalidOperationException("Не удалось создать контекст домена. Проверьте домен и OU DN.");

                    if (FindUser(ctx, sam) != null)
                        throw new InvalidOperationException($"Пользователь с логином '{sam}' уже существует.");

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
                var domain = SafeTrim(txtDomain.Text);
                var compName = SafeTrim(txtCompName.Text);
                if (string.IsNullOrWhiteSpace(compName)) throw new ArgumentException("Не указано имя компьютера.");

                if (cboOuComputers.SelectedItem is not OuItem ouItem)
                    throw new ArgumentException("Выберите OU для компьютера.");

                var containerDn = ouItem.DistinguishedName;
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

        // ===== Загрузка списков OU =====
        private void LoadOuListForUsers()
        {
            try
            {
                cboOuUsers.Items.Clear();
                foreach (var ou in EnumerateOuForUsers(SafeTrim(txtDomain.Text)))
                    cboOuUsers.Items.Add(ou);
                if (cboOuUsers.Items.Count > 0) cboOuUsers.SelectedIndex = 0;
                LogOk($"Загружено OU для пользователей: {cboOuUsers.Items.Count}");
            }
            catch (Exception ex)
            {
                LogErr("Не удалось загрузить OU для пользователей: " + ex.Message);
            }
        }
        private void LoadOuListForComputers()
        {
            try
            {
                cboOuComputers.Items.Clear();
                foreach (var ou in EnumerateOuForComputers(SafeTrim(txtDomain.Text)))
                    cboOuComputers.Items.Add(ou);
                if (cboOuComputers.Items.Count > 0) cboOuComputers.SelectedIndex = 0;
                LogOk($"Загружено OU для компьютеров: {cboOuComputers.Items.Count}");
            }
            catch (Exception ex)
            {
                LogErr("Не удалось загрузить OU для компьютеров: " + ex.Message);
            }
        }

        private IEnumerable<OuItem> EnumerateOuForUsers(string domainFqdn)
        {
            var (defaultNc, list) = EnumerateAllOus(domainFqdn);
            // Добавим стандартный контейнер CN=Users, если есть
            list.Add(new OuItem { Name = "CN=Users (стандартный)", DistinguishedName = "CN=Users," + defaultNc });
            return list.OrderBy(o => o.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
        }
        private IEnumerable<OuItem> EnumerateOuForComputers(string domainFqdn)
        {
            var (defaultNc, list) = EnumerateAllOus(domainFqdn);
            // Добавим стандартный контейнер CN=Computers, если есть
            list.Add(new OuItem { Name = "CN=Computers (стандартный)", DistinguishedName = "CN=Computers," + defaultNc });
            return list.OrderBy(o => o.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
        }

        // Единая функция поиска всех OU (organizationalUnit) в домене
        private (string defaultNc, List<OuItem> list) EnumerateAllOus(string domainFqdn)
        {
            // Получаем defaultNamingContext через RootDSE домена (если указан), иначе — текущего
            string defaultNc = null;
            try
            {
                DirectoryEntry rootDse;
                if (!string.IsNullOrWhiteSpace(domainFqdn))
                    rootDse = new DirectoryEntry($"LDAP://{domainFqdn}/RootDSE");
                else
                    rootDse = new DirectoryEntry("LDAP://RootDSE");

                defaultNc = rootDse.Properties["defaultNamingContext"]?.Value as string;
            }
            catch { }

            if (string.IsNullOrWhiteSpace(defaultNc))
                throw new InvalidOperationException("Не удалось определить defaultNamingContext. Проверьте подключение к домену.");

            var list = new List<OuItem>();

            try
            {
                using var root = new DirectoryEntry($"LDAP://{(string.IsNullOrWhiteSpace(domainFqdn) ? string.Empty : domainFqdn + "/")}{defaultNc}");
                using var ds = new DirectorySearcher(root)
                {
                    Filter = "(objectClass=organizationalUnit)",
                    SearchScope = SearchScope.Subtree,
                    PageSize = 500
                };
                ds.PropertiesToLoad.Add("name");
                ds.PropertiesToLoad.Add("distinguishedName");
                foreach (SearchResult r in ds.FindAll())
                {
                    var dn = r.Properties["distinguishedName"]?[0]?.ToString();
                    var name = r.Properties["name"]?[0]?.ToString();
                    if (!string.IsNullOrWhiteSpace(dn))
                        list.Add(new OuItem { Name = name ?? dn, DistinguishedName = dn });
                }
            }
            catch { }

            return (defaultNc, list);
        }

        // ===== Вспомогательные =====
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
        private static string SafeTrim(string s) => string.IsNullOrWhiteSpace(s) ? string.Empty : s.Trim();
        private void LogOk(string message) => txtLog.AppendText($"[{DateTime.Now:T}] ✅ {message}\r\n");
        private void LogErr(string message) => txtLog.AppendText($"[{DateTime.Now:T}] ❌ Ошибка: {message}\r\n");

        private class OuItem
        {
            public string Name { get; set; }
            public string DistinguishedName { get; set; }
            public override string ToString() => $"{Name} — {DistinguishedName}";
        }
    }
}