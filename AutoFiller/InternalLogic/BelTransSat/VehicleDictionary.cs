using DataTracker.Utility;

namespace AutoFiller.InternalLogic.BelTransSat;

public static class VehicleDictionary
{
    public static Dictionary<string, string> Dictionary = new Dictionary<string, string>();

    public static void LoadDictionary()
    {
        string dictionary = TxtHandler.ReadFile("dictionary.txt");
        dictionary = dictionary.Replace('\r', '=');
        string[] dictionarySplit = dictionary.Split('\n');
        for (int i = 0; i < dictionarySplit.Length; i++)
        {
            string[] pair = dictionarySplit[i].Split('=');
            Dictionary.Add(pair[0], pair[1]);
        }
        
        Logger.Log("Dictionary loaded.");
    }
}