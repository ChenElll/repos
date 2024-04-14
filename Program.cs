using System;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Do you want to delete the Outlook cache folder? (Y/N)");
        string input = Console.ReadLine();

        if (input.ToUpper() == "Y")
        {
            DeleteOutlookCache();
        }
        else
        {
            Console.WriteLine("Operation aborted.");
        }
    }

    static void DeleteOutlookCache()
    {
        string outlookCacheFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\Microsoft\Outlook";
        if (Directory.Exists(outlookCacheFolderPath))
        {
            Directory.Delete(outlookCacheFolderPath, true);
            Console.WriteLine("Outlook cache folder deleted successfully.");
        }
        else
        {
            Console.WriteLine("Outlook cache folder not found.");
        }
    }
}
