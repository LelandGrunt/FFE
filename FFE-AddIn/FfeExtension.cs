using System;
using System.IO;
using System.Threading;

namespace FFE
{
    public static class FfeExtension
    {
        public static void WriteToFile(this string content, string fileName = null, string subFolder = "")
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, FfeSetting.Default.LogFolder, subFolder);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            if (fileName != null)
            {
                // Avoid concurrent file (write) access.
                fileName = $"{Thread.CurrentThread.ManagedThreadId}_{fileName}";
            }
            else
            {
                fileName = Path.GetRandomFileName();
            }

            string filePath = Path.Combine(path, fileName);
            File.WriteAllText(filePath, content);
        }
    }
}