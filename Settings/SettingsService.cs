using System;
using System.IO;
using Newtonsoft.Json;

namespace ValveFlangeMulti.Settings
{
    public sealed class UserSettings
    {
        public int Version { get; set; } = 1;
        public string LastExcelPath { get; set; } = "";
    }

    /// <summary>
    /// Global settings persistence for ValveFlangeMulti.
    /// Stores last selected Excel path.
    /// </summary>
    public static class SettingsService
    {
        private static readonly string FolderPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "MTMDigCon", "ValveFlangeMulti");

        private const string GlobalFileName = "ValveFlangeMultiSettings.global.json";

        public static UserSettings LoadGlobal()
        {
            try
            {
                string filePath = Path.Combine(FolderPath, GlobalFileName);
                if (!File.Exists(filePath))
                    return new UserSettings();

                string json = File.ReadAllText(filePath);
                var settings = JsonConvert.DeserializeObject<UserSettings>(json);
                return settings ?? new UserSettings();
            }
            catch
            {
                return new UserSettings();
            }
        }

        public static void SaveGlobal(UserSettings settings)
        {
            try
            {
                if (!Directory.Exists(FolderPath))
                    Directory.CreateDirectory(FolderPath);

                string filePath = Path.Combine(FolderPath, GlobalFileName);
                string json = JsonConvert.SerializeObject(settings ?? new UserSettings(), Formatting.Indented);
                File.WriteAllText(filePath, json);
            }
            catch
            {
                // best-effort persistence; ignore IO errors
            }
        }
    }
}
