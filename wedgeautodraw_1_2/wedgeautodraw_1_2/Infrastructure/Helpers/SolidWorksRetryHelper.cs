namespace wedgeautodraw_1_2.Infrastructure.Helpers;

public static class SolidWorksRetryHelper
{
    public static T Retry<T>(Func<T> action, int maxAttempts = 3, int delayMs = 500, string context = "")
    {
        for (int attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                var result = action();
                if (result is not null) return result;
            }
            catch (Exception ex)
            {
                Logger.Warn($"[Attempt {attempt}] Failed in {context}: {ex.Message}");
                if (attempt == maxAttempts) throw;
            }

            Thread.Sleep(delayMs);
        }

        return default;
    }

    public static bool Retry(Func<bool> action, int maxAttempts = 3, int delayMs = 500, string context = "")
    {
        for (int attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                if (action()) return true;
            }
            catch (Exception ex)
            {
                Logger.Warn($"[Attempt {attempt}] Failed in {context}: {ex.Message}");
                if (attempt == maxAttempts) throw;
            }

            Thread.Sleep(delayMs);
        }

        return false;
    }
}
