using System.IO;
using System.Reflection;

namespace DataTracker.Utility
{
    public class TxtHandler
    {
        public static string ReadFile(string fileName)
        {
            string exeDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string configFilePath = Path.Combine(exeDirectory, fileName);

            if (File.Exists(configFilePath))
            {
                return File.ReadAllText(configFilePath);
            }

            Logger.Log(fileName + " failed to load!");
            return "";
        }
    }
}