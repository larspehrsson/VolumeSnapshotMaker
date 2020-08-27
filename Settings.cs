using System;
using System.Windows.Forms;
using Microsoft.Win32;

namespace SnapshotMaker
{
    public static class settings
    {
        private static readonly RegistryKey RK = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\SnapshotMaker");

        static settings()
        {
            Get();
        }

        public static string saveinterval { get; set; }
        public static string drivelist { get; set; }
        public static bool notification { get; set; }

        private static T GetSetting<T>(string setting)
        {
            var value = RK.GetValue(setting);
            if (typeof(T).FullName == "System.Boolean" && value == null)
                value = "False";

            if (value == null || value == "")
                return default(T);

            return (T)Convert.ChangeType(value, typeof(T));
        }

        private static void SetSetting<T>(string setting, T value)
        {
            if (value == null)
                RK.SetValue(setting, "");
            else
                RK.SetValue(setting, value);
        }

        public static void Get()
        {
            saveinterval = GetSetting<string>("saveinterval");
            drivelist = GetSetting<string>("drivelist");
            notification = GetSetting<bool>("notification");
        }

        public static void Save()
        {
            SetSetting("saveinterval", saveinterval);
            SetSetting("drivelist", drivelist);
            SetSetting("notification", notification);
        }
    }
}