// using System;
// using System.IO;
// using Microsoft.Win32;


// class Program
// {
//     static void Main(string[] args)
//     {
//         Console.WriteLine("Do you want to delete the Outlook cache folder? (Y/N)");
//         string input = Console.ReadLine();
//         string outlookCacheFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\Microsoft\Outlook";

//         if (input.ToUpper() == "Y")
//         {
//             DeleteOutlookCache(outlookCacheFolderPath);
//         }
//         else
//         {
//             Console.WriteLine("Operation aborted.");
//         }
//         Console.WriteLine("Do you want to Add a Local Pofile? (Y/N)");
//         string input1 = Console.ReadLine();
//         if (input.ToUpper() == "Y")
//         {
//             AddLocalPofile();
//         }
//         else
//         {
//             Console.WriteLine("Operation aborted.");
//         }
//     }

//     static void DeleteOutlookCache(string outlookCacheFolderPath)
//     {

//         if (Directory.Exists(outlookCacheFolderPath))
//         {
//             Directory.Delete(outlookCacheFolderPath, true);
//             Console.WriteLine("Outlook cache folder deleted successfully.");
//         }
//         else
//         {
//             Console.WriteLine("Outlook cache folder not found.");
//         }
//     }


//     static void AddLocalPofile()
//     {
//         try
//         {

//   // Specify the Outlook profile name and server settings
//             string profileName = "MyProfile";
//             string serverName = "mail.example.com";
//             string userName = "user@example.com";
//             string email = "user@example.com";

//             // Create registry keys for the new Outlook profile
//             using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\Outlook\Profiles\" + profileName))
//             {
//                 key.SetValue("DefaultProfile", profileName);
//                 key.SetValue("11", userName);
//                 key.SetValue("12", email);
//             }

//             // Display success message
//             Console.WriteLine("Local Outlook profile added successfully.");


//         }
//         catch (Exception ex)
//         {
//             Console.WriteLine("Error occurred while adding local Outlook profile: " + ex.Message);
//         }


//     }
// }






using System;
using System.IO;
using Microsoft.Win32;

class Program
{
    static void Main(string[] args)
    {
        // Define paths to Mail 32-bit configuration files and registry keys
        string mailAppDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Outlook");
        string mailRegistryKey = @"Software\Microsoft\Office\16.0\Outlook";

        // Call functions to delete configuration files and registry keys
        DeleteMailConfigFiles(mailAppDataFolder);
        DeleteMailRegistryKeys(mailRegistryKey);

        // Display success message
        Console.WriteLine("Mail 32-bit settings reset successfully.");
    }

    static void DeleteMailConfigFiles(string configFolder)
    {
        try
        {
            if (Directory.Exists(configFolder))
            {
                Console.WriteLine("Deleting Mail 32-bit configuration files.");
                Directory.Delete(configFolder, true);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while deleting Mail 32-bit configuration files: " + ex.Message);
        }
    }

    static void DeleteMailRegistryKeys(string registryKey)
    {
        try
        {
          
                using (RegistryKey regKey = Registry.CurrentUser.OpenSubKey(registryKey, true))
                {
                    if (regKey != null)
                    {
                        Console.WriteLine("Deleting registry key: " + registryKey);
                        regKey.DeleteSubKeyTree("Profiles");
                    }
                }
            
        }
        catch (Exception ex)
        {
         
            Console.WriteLine("An error occurred while deleting Mail 32-bit registry keys: " + ex.Message);
        }
    }
}
