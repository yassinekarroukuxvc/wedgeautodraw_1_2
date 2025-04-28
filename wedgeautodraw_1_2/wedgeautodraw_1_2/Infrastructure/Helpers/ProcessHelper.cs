using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                Console.WriteLine($"Killed SolidWorks process (PID: {process.Id})");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Could not kill SolidWorks process (PID: {process.Id}): {ex.Message}");
            }
        }
    }
}
