using System;
using System.Reflection;

namespace AD
{
    /// <summary>
    /// Единая точка правды о версии приложения.
    /// Версия берётся из сборки, но можно задать "человеческую" метку релиза.
    /// </summary>
    public static class AppVersion
    {
        // Укажи здесь номер релиза/канал (например, из CI): 
        // Меняй вручную или генерируй в пайплайне.
        public const string ReleaseLabel = "1.0.0"; // ← меняется при релизе

        public static Version AssemblyVersion =>
            Assembly.GetExecutingAssembly().GetName().Version ?? new Version(1, 0, 0, 0);

        public static string InformationalVersion
        {
            get
            {
                var info = Assembly.GetExecutingAssembly()
                                   .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
                return string.IsNullOrWhiteSpace(info) ? ReleaseLabel : info!;
            }
        }

        /// <summary>
        /// Полная строка версии, показ в UI.
        /// </summary>
        public static string FullVersion =>
            $"v{InformationalVersion} (asm {AssemblyVersion})";
    }
}
