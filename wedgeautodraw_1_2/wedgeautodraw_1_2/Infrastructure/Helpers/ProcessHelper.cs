using System.Diagnostics;

namespace wedgeautodraw_1_2.Infrastructure.Helpers;

public static class ProcessHelper
{
    public static void KillAllSolidWorksProcesses()
    {
        foreach (var process in Process.GetProcessesByName("SLDWORKS"))
        {
            try
            {
                process.Kill();
                process.WaitForExit();
                Logger.Success($"Killed SolidWorks process (PID: {process.Id})");
            }
            catch (Exception ex)
            {
                Logger.Warn($"Could not kill SolidWorks process (PID: {process.Id}): {ex.Message}");
            }
        }
    }
}
