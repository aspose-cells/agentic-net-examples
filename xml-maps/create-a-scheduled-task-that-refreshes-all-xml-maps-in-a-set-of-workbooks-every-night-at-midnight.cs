using System;
using System.IO;
using System.Timers;
using Aspose.Cells;

namespace XmlMapRefreshScheduler
{
    class Program
    {
        // Folder containing the workbooks to refresh
        private static readonly string WorkbooksFolder = @"C:\Workbooks";

        // Folder containing the XML source files for the maps
        private static readonly string XmlDataFolder = @"C:\XmlData";

        // Timer that triggers the refresh operation (System.Timers.Timer)
        private static System.Timers.Timer _refreshTimer;

        static void Main()
        {
            try
            {
                // Verify that the required folders exist
                if (!Directory.Exists(WorkbooksFolder))
                {
                    Console.WriteLine($"Workbooks folder not found: {WorkbooksFolder}");
                    return;
                }

                if (!Directory.Exists(XmlDataFolder))
                {
                    Console.WriteLine($"XML data folder not found: {XmlDataFolder}");
                    return;
                }

                // Calculate the interval until the next midnight
                DateTime now = DateTime.Now;
                DateTime nextMidnight = now.Date.AddDays(1);
                double initialDelay = (nextMidnight - now).TotalMilliseconds;

                // Set up the recurring 24‑hour timer (first trigger handled separately)
                _refreshTimer = new System.Timers.Timer
                {
                    AutoReset = true,
                    Interval = TimeSpan.FromDays(1).TotalMilliseconds,
                    Enabled = false
                };
                _refreshTimer.Elapsed += OnRefreshTimerElapsed;

                // One‑shot timer to start the first refresh at midnight
                var startTimer = new System.Timers.Timer(initialDelay) { AutoReset = false };
                startTimer.Elapsed += (s, e) =>
                {
                    _refreshTimer.Start();          // start the 24‑hour recurring timer
                    RefreshAllWorkbooks();          // run the first refresh at midnight
                    startTimer.Dispose();
                };
                startTimer.Start();

                // Keep the application running
                Console.WriteLine("XML map refresh scheduler started. Press Enter to exit.");
                Console.ReadLine();

                // Clean up timers on exit
                _refreshTimer?.Stop();
                _refreshTimer?.Dispose();
                startTimer?.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error in Main: {ex.Message}");
            }
        }

        // Event handler that runs each night at midnight
        private static void OnRefreshTimerElapsed(object sender, ElapsedEventArgs e)
        {
            RefreshAllWorkbooks();
        }

        // Refreshes all XML maps in every workbook found in WorkbooksFolder
        private static void RefreshAllWorkbooks()
        {
            try
            {
                // Get all Excel files in the target folder
                string[] workbookFiles = Directory.GetFiles(WorkbooksFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

                foreach (string workbookPath in workbookFiles)
                {
                    if (!File.Exists(workbookPath))
                    {
                        Console.WriteLine($"Workbook not found: {workbookPath}");
                        continue;
                    }

                    Workbook workbook = null;
                    try
                    {
                        // Load the workbook
                        workbook = new Workbook(workbookPath);
                    }
                    catch (Exception loadEx)
                    {
                        Console.WriteLine($"Failed to load workbook '{Path.GetFileName(workbookPath)}': {loadEx.Message}");
                        continue;
                    }

                    // Iterate through each XML map in the workbook
                    XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;
                    for (int i = 0; i < xmlMaps.Count; i++)
                    {
                        XmlMap map = xmlMaps[i];

                        // Assume the XML source file name matches the map name with .xml extension
                        string xmlFileName = map.Name + ".xml";
                        string xmlFullPath = Path.Combine(XmlDataFolder, xmlFileName);

                        if (File.Exists(xmlFullPath))
                        {
                            try
                            {
                                // Import the XML data into the first worksheet starting at cell A1 (row 0, column 0)
                                workbook.ImportXml(xmlFullPath, workbook.Worksheets[0].Name, 0, 0);
                            }
                            catch (Exception importEx)
                            {
                                Console.WriteLine($"Error importing XML for map '{map.Name}' in workbook '{Path.GetFileName(workbookPath)}': {importEx.Message}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"XML source file not found for map '{map.Name}' in workbook '{Path.GetFileName(workbookPath)}'.");
                        }
                    }

                    try
                    {
                        // Save the workbook (overwrite the original file)
                        workbook.Save(workbookPath);
                        Console.WriteLine($"Refreshed XML maps in workbook: {Path.GetFileName(workbookPath)}");
                    }
                    catch (Exception saveEx)
                    {
                        Console.WriteLine($"Failed to save workbook '{Path.GetFileName(workbookPath)}': {saveEx.Message}");
                    }
                    finally
                    {
                        workbook?.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during XML map refresh: {ex.Message}");
            }
        }
    }
}