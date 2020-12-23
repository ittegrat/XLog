using System;
using System.Collections;
using System.Configuration;

namespace XLog
{
  internal static class Configuration
  {

    const string CONFIG_SECTION = "XLog";

    public static bool ArchiveOldFileOnStartup { get; }
    public static bool DeleteOldFileOnStartup { get; }

    public static string DefaultLayout { get; }
    public static string DefaultNumberSuffix { get; }
    public static string DefaultSuffix { get; }

    static Configuration() {
      Hashtable section = null;
      try {
        ConfigurationManager.RefreshSection(CONFIG_SECTION);
        section = (Hashtable)ConfigurationManager.GetSection(CONFIG_SECTION);
      }
      catch (Exception) {
        //** log error message
      }

      T GetValue<T>(string key, T @default) {
        if (section != null && section.ContainsKey(key))
          return (T)Convert.ChangeType(section[key], typeof(T));
        return @default;
      }

      ArchiveOldFileOnStartup = GetValue("archive.oldfiles", true);
      DeleteOldFileOnStartup = GetValue("delete.oldfiles", true);

      DefaultLayout = GetValue("default.layout", "${longdate}|${level:uppercase=true}|${message}");
      DefaultNumberSuffix = GetValue("default.number.suffix", ".{###}");
      DefaultSuffix = GetValue("default.suffix", ".log");

    }

  }
}
