using System;
using System.IO;

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
            string filePath = Path.Combine(path, fileName ?? Path.GetRandomFileName());
            File.WriteAllText(filePath, content);
        }
    }
}