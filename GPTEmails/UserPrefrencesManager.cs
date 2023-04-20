using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;
using System.Windows.Forms;

namespace GPTEmails
{

    internal class UserPreferences
    {
        public string selectedTemplate { get; set; }
        public string selectedSignature { get; set; }
        // Add other preferences as needed
    }

    internal class UserPrefrencesManager
    {

        private static readonly string PreferencesFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "YourAppName",
        "UserPreferences.json");

        public static void SaveUserPreferences(UserPreferences preferences)
        {
            string json = JsonConvert.SerializeObject(preferences, Formatting.Indented);

            Directory.CreateDirectory(Path.GetDirectoryName(PreferencesFilePath));
            File.WriteAllText(PreferencesFilePath, json);
        }

        public static UserPreferences LoadUserPreferences()
        {
            if (!File.Exists(PreferencesFilePath))
            {
                return new UserPreferences(); // Return default preferences if the file doesn't exist
            }

            string json = File.ReadAllText(PreferencesFilePath);
            UserPreferences preferences = JsonConvert.DeserializeObject<UserPreferences>(json);
            return preferences;
        }

    }
}
