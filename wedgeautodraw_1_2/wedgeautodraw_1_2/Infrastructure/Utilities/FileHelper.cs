namespace wedgeautodraw_1_2.Infrastructure.Utilities;

public static class FileHelper
{
    public static void CopyTemplateFile(string sourcePath, string destinationPath)
    {
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException("Template file not found", sourcePath);

        if (File.Exists(destinationPath))
            File.Delete(destinationPath);

        File.Copy(sourcePath, destinationPath);
    }
}
