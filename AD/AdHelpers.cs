using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace AD
{
    // ===== Модели для узлов дерева =====
    public class OuNodeTag
    {
        public string Name { get; set; }
        public string DistinguishedName { get; set; }
        public NodeKind Kind { get; set; }
    }

    public class AccountNodeTag
    {
        public string Name { get; set; }
        public string SamAccountName { get; set; }
        public string DistinguishedName { get; set; }
        public NodeKind Kind { get; set; } // User / Computer
    }

    public enum NodeKind
    {
        DomainRoot,
        OrganizationalUnit,
        Container,
        User,
        Computer
    }

    // ===== Утилиты AD/строки/пароли/поиск =====
    public static class AdUtils
    {
        public static string GetDefaultUpnSuffix()
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

        public static PrincipalContext BuildPrincipalContext(string domainFqdn, string containerDn)
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
            catch
            {
                return null;
            }
        }

        public static UserPrincipal FindUser(PrincipalContext ctx, string sam)
        {
            try { return UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, sam); } catch { return null; }
        }

        public static ComputerPrincipal FindComputer(PrincipalContext ctx, string compName)
        {
            try { return ComputerPrincipal.FindByIdentity(ctx, IdentityType.Name, compName); } catch { return null; }
        }

        public static string BuildLdapPath(string domainFqdn, string dn)
        {
            var prefix = string.IsNullOrWhiteSpace(domainFqdn) ? "LDAP://" : $"LDAP://{domainFqdn}/";
            return prefix + dn;
        }

        public static bool IsObjectClass(DirectoryEntry de, string className)
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

        public static bool NameEquals(DirectoryEntry de, string expected)
        {
            try
            {
                var n = de.Properties["name"]?.Value?.ToString();
                return string.Equals("name=" + n, expected, StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(n, expected.Split('=')[1], StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        public static string SafeProp(DirectoryEntry de, string prop)
        {
            try { return de.Properties[prop]?.Value?.ToString() ?? string.Empty; } catch { return string.Empty; }
        }

        public static string EscapeLdapValue(string value)
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

        public static string SafeTrim(string s) => string.IsNullOrWhiteSpace(s) ? string.Empty : s.Trim();

        // ===== Автогенерация логина (ФИО → sAMAccountName) =====
        public static string BuildBestGuessSam(string given, string surname)
        {
            var g = TransliterateRuToEn(given);
            var s = TransliterateRuToEn(surname);

            g = SlugifyAscii(g);
            s = SlugifyAscii(s);

            var initial = g.Length > 0 ? g[0].ToString() : "";
            var cand = (initial.Length > 0 ? $"{initial}.{s}" : s);

            cand = TrimSamToMax(cand).Trim('.');
            if (string.IsNullOrWhiteSpace(cand)) cand = "user";

            return cand.ToLowerInvariant();
        }

        public static string GenerateUniqueSam(PrincipalContext ctx, string baseSam, string given, string surname, int maxLen, Action<string> log)
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

        // ===== Транслитерация RU→EN (упрощённая) =====
        public static string TransliterateRuToEn(string src)
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
        public static string SlugifyAscii(string s)
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

        public static string TrimSamToMax(string s, int max = 20)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return s.Length <= max ? s : s.Substring(0, max);
        }

        // ===== Генерация безопасного пароля =====
        public static string GenerateSecurePassword(int length = 16)
        {
            if (length < 12) length = 12;

            const string U = "ABCDEFGHJKLMNPQRSTUVWXYZ";
            const string L = "abcdefghijkmnopqrstuvwxyz";
            const string D = "23456789";
            const string S = "!@#$%^&*()-_=+[]{};:,.?/";

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

        // ===== Быстрые LDAP-методы (для ленивой подгрузки) =====
        public static List<OuNodeTag> FetchChildContainers(string domainFqdn, string parentDn)
        {
            var result = new List<OuNodeTag>();

            using var parent = new DirectoryEntry(
                BuildLdapPath(domainFqdn, parentDn),
                null, null, AuthenticationTypes.Secure // явная безопасная аутентификация
            );

            // 1) Попытка через DirectorySearcher (быстро)
            using (var ds = new DirectorySearcher(parent))
            {
                ds.SearchScope = SearchScope.OneLevel;
                ds.PageSize = 1000;
                ds.PropertyNamesOnly = true;
                ds.ReferralChasing = ReferralChasingOption.All;
                // Более совместимый фильтр для OU и контейнеров
                ds.Filter =
                    "(|" +
                      "(&(objectClass=organizationalUnit)(objectCategory=organizationalUnit))" +
                      "(&(objectClass=container)(!(objectClass=computer)) )" + // «чистые» контейнеры
                    ")";

                ds.PropertiesToLoad.Add("name");
                ds.PropertiesToLoad.Add("distinguishedName");
                ds.PropertiesToLoad.Add("objectClass");

                try
                {
                    foreach (SearchResult r in ds.FindAll())
                    {
                        var dn = r.Properties["distinguishedName"]?.Count > 0 ? r.Properties["distinguishedName"][0]?.ToString() : null;
                        var name = r.Properties["name"]?.Count > 0 ? r.Properties["name"][0]?.ToString() : null;
                        if (string.IsNullOrWhiteSpace(dn) || string.IsNullOrWhiteSpace(name)) continue;

                        var kind = NodeKind.Container;
                        var oc = r.Properties["objectClass"];
                        if (oc != null)
                        {
                            foreach (var v in oc)
                            {
                                if (string.Equals(v?.ToString(), "organizationalUnit", StringComparison.OrdinalIgnoreCase))
                                {
                                    kind = NodeKind.OrganizationalUnit; break;
                                }
                            }
                        }

                        result.Add(new OuNodeTag { Name = name, DistinguishedName = dn, Kind = kind });
                    }
                }
                catch
                {
                    // Переходим к резервному способу ниже
                }
            }

            // 2) РЕЗЕРВ: прямой обход Children (работает там, где DS «молчит»)
            if (result.Count == 0)
            {
                foreach (DirectoryEntry ch in parent.Children)
                {
                    try
                    {
                        var cls = ch.SchemaClassName; // быстрее, чем objectClass
                        if (!string.Equals(cls, "organizationalUnit", StringComparison.OrdinalIgnoreCase) &&
                            !string.Equals(cls, "container", StringComparison.OrdinalIgnoreCase))
                            continue;

                        var dn = SafeProp(ch, "distinguishedName");
                        var name = SafeProp(ch, "name");
                        if (string.IsNullOrWhiteSpace(dn) || string.IsNullOrWhiteSpace(name)) continue;

                        var kind = string.Equals(cls, "organizationalUnit", StringComparison.OrdinalIgnoreCase)
                                   ? NodeKind.OrganizationalUnit : NodeKind.Container;

                        result.Add(new OuNodeTag { Name = name, DistinguishedName = dn, Kind = kind });
                    }
                    catch { /* пропускаем проблемные ветки */ }
                }
            }

            return result;
        }

        public static OuNodeTag TryGetNamedContainer(string domainFqdn, string baseDn, string cnName)
        {
            using var baseDe = new DirectoryEntry(
                BuildLdapPath(domainFqdn, baseDn),
                null, null, AuthenticationTypes.Secure
            );
            using var ds = new DirectorySearcher(baseDe)
            {
                SearchScope = SearchScope.OneLevel,
                PageSize = 50,
                PropertyNamesOnly = true,
                ReferralChasing = ReferralChasingOption.All,
                Filter = $"(&(objectClass=container)(name={EscapeLdapValue(cnName.Split('=')[1])}))"
            };

            ds.PropertiesToLoad.Add("name");
            ds.PropertiesToLoad.Add("distinguishedName");

            try
            {
                var res = ds.FindOne();
                if (res != null)
                {
                    var dn = res.Properties["distinguishedName"]?.Count > 0 ? res.Properties["distinguishedName"][0]?.ToString() : null;
                    var name = res.Properties["name"]?.Count > 0 ? res.Properties["name"][0]?.ToString() : null;
                    if (!string.IsNullOrWhiteSpace(dn))
                        return new OuNodeTag { Name = name ?? dn, DistinguishedName = dn, Kind = NodeKind.Container };
                }
            }
            catch
            {
                // Резерв через Children
                foreach (DirectoryEntry ch in baseDe.Children)
                {
                    try
                    {
                        if (!string.Equals(ch.SchemaClassName, "container", StringComparison.OrdinalIgnoreCase)) continue;
                        var n = SafeProp(ch, "name");
                        if (!string.Equals(n, cnName.Split('=')[1], StringComparison.OrdinalIgnoreCase)) continue;

                        var dn = SafeProp(ch, "distinguishedName");
                        return new OuNodeTag { Name = n, DistinguishedName = dn, Kind = NodeKind.Container };
                    }
                    catch { }
                }
            }
            return null;
        }

        public static List<AccountNodeTag> FetchAccounts(string domainFqdn, string containerDn, bool includeUsers, bool includeComputers)
        {
            var list = new List<AccountNodeTag>();
            using var de = new DirectoryEntry(BuildLdapPath(domainFqdn, containerDn));
            using var ds = new DirectorySearcher(de)
            {
                SearchScope = SearchScope.OneLevel,
                PageSize = 1000,
                PropertyNamesOnly = false
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

                    list.Add(new AccountNodeTag
                    {
                        Name = shown,
                        SamAccountName = sam,
                        DistinguishedName = dn,
                        Kind = NodeKind.User
                    });
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

                    list.Add(new AccountNodeTag
                    {
                        Name = shown,
                        SamAccountName = sam,
                        DistinguishedName = dn,
                        Kind = NodeKind.Computer
                    });
                }
            }

            return list;
        }
    }
}
