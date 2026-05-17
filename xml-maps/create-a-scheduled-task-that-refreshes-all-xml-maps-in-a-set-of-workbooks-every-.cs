using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace XmlMapRefreshScheduler
{
    class Program
    {
        // List of workbook files to process
        private static readonly List<string> WorkbookFiles = new List<string>
        {
            @"C:\Workbooks\Book1.xlsx",
            @"C:\Workbooks\Book2.xlsx"
            // add more paths as needed
        };

        static void Main()
        {
            // Calculate time until next midnight
            DateTime now = DateTime.Now;
            DateTime nextMidnight = now.Date.AddDays(1);
            TimeSpan dueTime = nextMidnight - now;

            // Set up a timer that fires at midnight and then every 24 hours
            Timer timer = new Timer(RefreshAllWorkbooks, null, dueTime, TimeSpan.FromHours(24));

            // Keep the application running
            Console.WriteLine("XML map refresh scheduler started. Press Enter to exit.");
            Console.ReadLine();
        }

        // Callback executed by the timer
        private static void RefreshAllWorkbooks(object state)
        {
            Console.WriteLine($"Refresh started at {DateTime.Now}");

            foreach (string filePath in WorkbookFiles)
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (uses the provided load rule)
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        // Iterate through all XML maps in the workbook
                        foreach (XmlMap xmlMap in workbook.Worksheets.XmlMaps)
                        {
                            // Determine the source XML file for the map.
                            // This example assumes the XML file name matches the map name with .xml extension.
                            // Adjust the logic as needed for your environment.
                            string xmlSourcePath = Path.Combine(Path.GetDirectoryName(filePath) ?? string.Empty, $"{xmlMap.Name}.xml");

                            if (File.Exists(xmlSourcePath))
                            {
                                // Import the XML data into the first worksheet at cell A1.
                                // This effectively refreshes the XML map with the latest data.
                                workbook.ImportXml(xmlSourcePath, workbook.Worksheets[0].Name, 0, 0);
                                Console.WriteLine($"Refreshed XML map '{xmlMap.Name}' in '{Path.GetFileName(filePath)}' using '{Path.GetFileName(xmlSourcePath)}'.");
                            }
                            else
                            {
                                Console.WriteLine($"XML source not found for map '{xmlMap.Name}' in workbook '{Path.GetFileName(filePath)}'. Expected at '{xmlSourcePath}'.");
                            }
                        }

                        // Save the workbook (uses the provided save rule)
                        workbook.Save(filePath);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine($"Refresh completed at {DateTime.Now}");
        }
    }
}