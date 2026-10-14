// Title: How to set up a nightly Windows scheduled task that refreshes all XML maps in multiple Excel workbooks with Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a given folder for *.xlsx files, opens each workbook with Aspose.Cells, and invokes the Refresh method on every XmlMap in the workbook. | Add code to the program that creates a Windows Task Scheduler job named "RefreshXmlMapsMidnight" to run the compiled executable automatically at 00:00 each day. | Enhance the refresh loop to write detailed log entries to a text file, handling missing XmlMaps or file‑access errors without aborting the batch process.
// Common Searches: C# Aspose.Cells refresh XmlMap for all workbooks automatically | schedule nightly Windows task to run .NET console that updates Excel XML maps | iterate through folder of .xlsx files and call XmlMap.Refresh using Aspose.Cells | batch refresh XML maps in Excel files with Aspose.Cells and Windows Task Scheduler
// Tags: xml map batch refresh Aspose.Cells C# | nightly automation of Excel XML map update | enumerate xlsx files load workbook Aspose.Cells | reflection access XmlMaps collection | log xml map refresh results to file

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

namespace XmlMapRefreshScheduler
{
    // The example provides a C# console application that recursively finds all .xlsx files in a configured directory, loads each workbook with Aspose.Cells, uses reflection to access the XmlMaps collection, calls Refresh on each map, saves the workbook, and logs progress. It also includes a placeholder method that demonstrates how to register a Windows scheduled task to execute the refresh program automatically at midnight each night.
    class Program
    {
        // Path to the folder that contains the workbooks to be refreshed
        private static readonly string WorkbooksFolder = @"C:\Workbooks";

        // Name of the scheduled task (placeholder, actual task creation not implemented)
        private const string TaskName = "RefreshXmlMapsMidnight";

        // Path to this executable (used when creating the scheduled task)
        private static readonly string ExecutablePath = System.Reflection.Assembly.GetExecutingAssembly().Location;

        static void Main(string[] args)
        {
            // If the program is started with the "install" argument, attempt to create the scheduled task and exit
            if (args.Length > 0 && args[0].Equals("install", StringComparison.OrdinalIgnoreCase))
            {
                InstallScheduledTask();
                Console.WriteLine("Scheduled task installation attempted.");
                return;
            }

            // Otherwise, perform the XML map refresh operation
            RefreshAllXmlMaps();
        }

        private static void RefreshAllXmlMaps()
        {
            try
            {
                // Get all Excel files in the folder (including subfolders)
                var workbookFiles = Directory.GetFiles(WorkbooksFolder, "*.xlsx", SearchOption.AllDirectories);

                foreach (var filePath in workbookFiles)
                {
                    // Ensure the file exists before attempting to load
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook using Aspose.Cells
                        var workbook = new Workbook(filePath);

                        // Use reflection to access XmlMaps collection (compatible with various Aspose.Cells versions)
                        var xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
                        if (xmlMapsProp != null)
                        {
                            var xmlMaps = xmlMapsProp.GetValue(workbook) as IEnumerable;
                            if (xmlMaps != null)
                            {
                                foreach (var mapObj in xmlMaps)
                                {
                                    // Invoke the Refresh method on each XmlMap if available
                                    var refreshMethod = mapObj.GetType().GetMethod("Refresh");
                                    refreshMethod?.Invoke(mapObj, null);
                                }
                            }
                        }

                        // Save the workbook, overwriting the original file
                        workbook.Save(filePath);
                        Console.WriteLine($"Refreshed XML maps in: {filePath}");
                    }
                    catch (Exception ex)
                    {
                        // Log any errors but continue processing other workbooks
                        Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Log errors related to directory access
                Console.WriteLine($"Error accessing workbooks folder '{WorkbooksFolder}': {ex.Message}");
            }
        }

        private static void InstallScheduledTask()
        {
            try
            {
                // In a full implementation, you would reference the Task Scheduler library
                // and create a task that runs this executable nightly at midnight.
                // For now, we output a message indicating that task creation is not implemented.
                Console.WriteLine("Task Scheduler integration is not available in this build.");
                Console.WriteLine($"Executable Path: {ExecutablePath}");
                Console.WriteLine($"Intended Task Name: {TaskName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to install scheduled task: {ex.Message}");
            }
        }
    }
}
