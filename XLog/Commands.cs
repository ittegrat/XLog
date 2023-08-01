using ExcelDna.Integration;

namespace XLog
{

  [ExcelCommand(Prefix = "LogDisplay.")]
  public static class LogDisplayCommands
  {
    public static void Clear() {
      ExcelDna.Logging.LogDisplay.Clear();
    }
    public static void Hide() {
      ExcelDna.Logging.LogDisplay.Hide();
    }
    public static void Show() {
      ExcelDna.Logging.LogDisplay.Show();
    }
  }

}
