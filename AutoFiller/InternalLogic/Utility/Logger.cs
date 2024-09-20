using System.Collections.ObjectModel;

namespace DataTracker.Utility
{
    public class Logger
    {
        public static ObservableCollection<string> Logs { get; private set; } = new ObservableCollection<string>();
        public static void Log(string message)
        {
            string logEntry = $"[{DateTime.Now}] {message}";
            //Console.WriteLine($"[{DateTime.Now}] {message}");
            Logs.Add(logEntry);
        }
    }
}