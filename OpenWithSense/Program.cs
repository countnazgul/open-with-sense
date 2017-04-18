using System;
using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using Qlik.Engine;
//using Qlik.Sense;
using System.Diagnostics;
using System.Net;
using Qlik.Sense.Client;
using System.IO;

namespace OpenWithSense
{
    class Program
    {
        static void Main(string[] args)
        {

#if DEBUG
            string currentExeName = "OpenWithSense.exe";
#else
            string currentExeName = System.AppDomain.CurrentDomain.FriendlyName;
#endif



            IHub hub;
            ILocation location = Qlik.Engine.Location.FromUri("ws://127.0.0.1:4848");
            location.AsDirectConnectionToPersonalEdition();
            string version = "";
            string welcome =  "Enter the needed number or enter 0 to exit in any menu";

            // New app name prefix
            string tempAppPrefix = "Temp App - ";
            string userName = System.Environment.UserName;
            string sensePath = @"C:\Users\" + userName + @"\AppData\Local\Programs\Qlik\Sense\QlikSense.exe";
            string currentFolder = System.AppDomain.CurrentDomain.BaseDirectory;
            string keyPath = "Software\\Classes\\*\\Shell\\OpenInSense";

            // Possible file extensions       
            List<string> fileExtensions = new List<string>();
            fileExtensions.Add(".csv");
            fileExtensions.Add(".qvd");
            fileExtensions.Add(".qvf");

            //var enviromentPath1 = System.Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.User);

            // Check if Sense is running before everything else. If not - exit
            try
            {
                using (hub = location.Hub())
                {
                    version = hub.ProductVersion();

                    if (version.Substring(0,3) != "3.2")
                    {
                        Console.WriteLine("I'm sory but only QS version 3.2.x is supported at the moment. Press Enter to close.");
                        Console.ReadLine();
                        Environment.Exit(0);
                    }
                }
            } catch (Exception ex)
            {
                Console.Clear();
                Console.WriteLine("Qlik Sense Desktop is not running. Please start it and try again.");
                Console.WriteLine("Bye!");
                Environment.Exit(0);
            }


            // Check if the app is started from the right click or directly from command line
            if (args.Length > 0)
            {
                string loadedExtension = Path.GetExtension(args[0]).ToLower();
                string loadedFileName = Path.GetFileName(args[0]);

                bool possible = false;
                for(var i = 0; i < fileExtensions.Count; i++)
                {
                    if(fileExtensions[i] == loadedExtension)
                    {
                        possible = true;
                    }
                }

                // if the selected extension is in the list - csv, qvd and qvf
                if (possible == true)
                {
                    using (hub = location.Hub())
                    {
                        // if its data file - csv or qvd
                        if (loadedExtension != ".qvf")
                        {
                            // Create new app with prefix name and timestamp
                            // Open the app; GetSript; Add the new script; Reload the app; add summary sheet; save the app and open the app inside browser on the summary sheet in edit mode

                            //Console.WriteLine(hub.ProductVersion());
                            string currentTime = DateTime.Now.ToString("yyyyMMddHHmmssffff");
                            string appName = tempAppPrefix + "" + currentTime;
                            Qlik.Engine.CreateAppResult appResult = hub.CreateApp(appName);
                            Console.WriteLine("App is created --> '" + appName + "'");
                            IApp app = hub.OpenApp(appName);
                            Console.WriteLine("App is open ...");
                            string script = app.GetScript();
                            app.SetScript(script + Environment.NewLine + Environment.NewLine + "Load 1 as test AutoGenerate(1);");
                            Console.WriteLine("New script is added ...");
                            Console.WriteLine("Reload started...");
                            app.DoReload();
                            Console.WriteLine("App is reloaded ...");


                            ISheet mySheet = app.CreateSheet("Summary");
                            using (mySheet.SuspendedLayout)
                            {
                                mySheet.Properties.MetaDef.Title = "Summary";
                                mySheet.Properties.Rank = 0;
                            }

                            Console.WriteLine("New sheet is added ...");
                            Console.WriteLine("App is saving ...");
                            app.DoSave();
                            Console.WriteLine("App is saved.");
                            Console.WriteLine("All is done!");
                            Process.Start("http://localhost:4848/sense/app/" + WebUtility.UrlEncode(appResult.AppId).Replace("+", "%20") + "/sheet/Summary/state/edit");
                            Console.Clear();
                            Console.WriteLine("Bye!");
                            Environment.Exit(0);
                        }
                        // if its qvf
                        else
                        {
                            string destinationPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Qlik\\Sense\\Apps\\" + loadedFileName;
                            
                            // check if the file already exists in QS Apps folder
                            // if not - copy it
                            // if yes - ask the user to confirm if its ok to overwrite
                            if (!File.Exists(destinationPath))
                            {
                                CopyQVF(args[0], destinationPath, false);
                            } else
                            {
                                Console.Write("The app already exists. Overwrite? (y/n) ");
                                string result = Console.ReadLine().ToLower();

                                if(result == "y")
                                {
                                    CopyQVF(args[0], destinationPath, true);
                                } else
                                {
                                    Console.Clear();
                                    Console.WriteLine("Nothing to do then");
                                    Console.WriteLine("Bye!");
                                    Environment.Exit(0);
                                }
                            }

                        }
                    }
                } else
                {
                    Console.Clear();
                    Console.WriteLine("I'm sorry but at the moment only qvf, qvd and csv files are supported (Press Enter to exit)");
                    Console.ReadLine();
                }

            } else
            {
                //Print the possible options
                var options = PrintOptions(welcome + "", new string[] {
                                                                        "1. Clear all temporary apps",
                                                                        "2. Add action on right click",
                                                                        "3. Remove the right click action",
                                                                        "4. View documentation (in GitHub)",
                                                                        "0. Exit" }, true);
                switch(options)
                {
                    // Exit
                    case 0:
                        {
                            Console.Clear();
                            Console.WriteLine("Bye!");
                            Environment.Exit(0);
                            break;
                        }
                    // Get all apps and filter only the ones that have app name prefix in it.
                    // List all the prefixed ones and ask the user ro confirm deletion of all.
                    // If no apps were found - exit
                    case 1:
                        using (hub = location.Hub())
                        {
                            int tempAppsCount = 0;
                            List<string> tempAppNames = new List<string>();
                            List<string> tempAppIds = new List<string>();
                            foreach (IAppIdentifier appIdentifier in location.GetAppIdentifiers())
                            {
                                if (appIdentifier.AppName.IndexOf(tempAppPrefix) > -1)
                                {
                                    tempAppsCount++;
                                    tempAppNames.Add(appIdentifier.AppName);
                                    tempAppIds.Add(appIdentifier.AppId);
                                }
                            }

                            if(tempAppsCount > 0)
                            {
                                Console.Clear();
                                Console.WriteLine( tempAppsCount + " temp apps found");
                                Console.WriteLine("---------------------------------");
                                for(int i = 0; i < tempAppNames.Count; i++)
                                {
                                    Console.WriteLine(tempAppNames[i]);
                                }
                                Console.WriteLine("---------------------------------");
                                Console.WriteLine();
                                Console.Write("Delete all? (y/n)");
                                string response = Console.ReadLine();

                                if(response.ToLower() == "y")
                                {
                                    for (int i = 0; i < tempAppIds.Count; i++)
                                    {
                                        hub.DestroyApp(tempAppIds[i]);
                                    }

                                    Console.Clear();
                                    Console.WriteLine(tempAppsCount + " app(s) deleted.");
                                    Console.WriteLine("Bye!");
                                    Environment.Exit(0);
                                } else
                                {
                                    Console.Clear();
                                    Console.WriteLine("Bye!");
                                    Environment.Exit(0);
                                }
                            } else
                            {
                                Console.Clear();
                                Console.WriteLine("0 Temp apps found");
                                Console.WriteLine("Bye!");
                                Environment.Exit(0);
                            }                            
                        }
                            break;

                    // Write to registry the values and keys needed for the context menu
                    // Also the folder is added to PATH variable so the app is accessible from everywhere
                    // This is madatory step after the app is started for the first time
                    case 2:
                        {
                            Console.Clear();
                            Microsoft.Win32.RegistryKey key;
                            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(keyPath);
                            key.SetValue("", "Load in Qlik Sense");
                            key.SetValue("Icon", sensePath);
                            key.Close();
                            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(keyPath + "\\command");
                            key.SetValue("", currentFolder + currentExeName + " %1");

                            var enviromentPath = System.Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.User);
                            System.Environment.SetEnvironmentVariable("PATH", enviromentPath + ";" + currentFolder, EnvironmentVariableTarget.User);

                            Console.WriteLine("Right click option is added");
                            Console.WriteLine("Folder is added to Path variable - just type 'OpenWithSense' in command prompth to see the options again");
                            Console.WriteLine("Bye!");
                            Environment.Exit(0);
                            break;
                        }
                    // Delete the registry and the portion of PATH value
                    // Kinda uninstall option
                    case 3:
                        {
                            Console.Clear();
                            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser;
                            key.DeleteSubKeyTree(keyPath,false);
                            key.Close();

                            var enviromentPath = System.Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.User);
                            enviromentPath = enviromentPath.Replace(currentFolder, "");
                            System.Environment.SetEnvironmentVariable("PATH", enviromentPath, EnvironmentVariableTarget.User);

                            Console.WriteLine("Right click option is removed");
                            Console.WriteLine("Path variable is changed and the base folder is removed from it");
                            Console.WriteLine("Bye!");
                            Environment.Exit(0);
                            break;
                        }
                    // Opens the browser to the GitHub page of this app repo
                    case 4:
                        {
                            Console.Clear();
                            Process.Start("https://github.com/countnazgul/open-with-sense/blob/master/README.md");
                            Environment.Exit(0);
                            break;
                        }
                    // Exit at any non valid selection
                    default:
                        {
                            Console.Clear();
                            Console.WriteLine("Invalid selection.");
                            Console.WriteLine("Bye!");
                            Environment.Exit(0);
                            break;
                        }
                }
            }
        }

        // Function that prints the available options
        static public int PrintOptions(string header, string[] options, bool clear)
        {
            if (clear == true)
            {
                Console.Clear();
            }
            Console.WriteLine(header);
            Console.WriteLine("");
            


            for (var i = 0; i < options.Length; i++)
            {
                Console.WriteLine(options[i]);
            }

            Console.WriteLine("");
            Console.Write("Pick a number: ");

            int a = 0;

            try
            {
                a = int.Parse(Console.ReadLine());
            }
            catch (Exception ex)
            {
                
            }

            return a;
        }

        // Function that actually copy the selected qvf file to QS Apps folder
        static public void CopyQVF(string source, string destination, bool overwrite)
        {
            Console.WriteLine("File copy started ...");
            Process.Start("http://localhost:4848/sense/app/" + WebUtility.UrlEncode(destination).Replace("+", "%20"));
            File.Copy(source, destination, overwrite);
            Console.Clear();
            Console.WriteLine("File copy has finished");
            Console.WriteLine("Bye!");
            Environment.Exit(0);
        }
    }
}
