// Title: Create a nightly Windows Task Scheduler job that refreshes all XML maps in Excel workbooks with Aspose.Cells for .NET
// AI Prompts: Generate C# code that enumerates all .xlsx files in a folder, loads each workbook with Aspose.Cells, refreshes every XmlMap, and saves the file. | Write a method that registers a Windows Task Scheduler task to run the refresh program automatically at 00:00 each day. | Add comprehensive try‑catch blocks and logging to capture failures when refreshing XmlMaps or creating the scheduled task.
// Common Searches: how to refresh xml maps in multiple Excel files using Aspose.Cells C# | schedule a .NET console application to run at midnight with Windows Task Scheduler | batch process xmlmap collection in workbooks Aspose.Cells example | c# program to iterate over workbook.XmlMaps and call Refresh | automate nightly Excel XML map update on Windows server
// Tags: Aspose.Cells refresh XmlMap batch | C# Windows Task Scheduler automation | process multiple .xlsx workbooks programmatically | nightly XML map refresh for Excel | dynamic workbook loading Aspose.Cells

using System;
using System.IO;
using System.Diagnostics;
using Aspose.Cells;

namespace XmlMapRefresher
{
    // A C# console utility that can create a Windows Task Scheduler task to run at midnight, or when executed, scans a configured directory for .xlsx files, loads each workbook with Aspose.Cells, refreshes all XmlMaps, saves the changes, and logs any errors encountered.
    class Program
    {
        // Directory containing the workbooks to process
        private const string WorkbooksFolder = @"C:\Workbooks";

        // Path to the compiled executable (this program)
        private static readonly string ExecutablePath = System.Reflection.Assembly.GetExecutingAssembly().Location;

        static void Main(string[] args)
        {
            try
            {
                // If the program is called with the argument "setup", create the scheduled task.
                // Otherwise, perform the XML map refresh operation.
                if (args.Length > 0 && args[0].Equals("setup", StringComparison.OrdinalIgnoreCase))
                {
                    CreateMidnightTask();
                    Console.WriteLine("Scheduled task created successfully.");
                }
                else
                {
                    RefreshAllXmlMaps();
                    Console.WriteLine("All XML maps refreshed.");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Refreshes all XML maps in every workbook found in WorkbooksFolder.
        private static void RefreshAllXmlMaps()
        {
            // Get all Excel files in the target folder (including subfolders)
            string[] files = Directory.GetFiles(WorkbooksFolder, "*.xlsx", SearchOption.AllDirectories);

            foreach (string filePath in files)
            {
                try
                {
                    // Ensure the file exists before attempting to load
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found, skipping: {filePath}");
                        continue;
                    }

                    // Load the workbook using dynamic to avoid compile‑time binding issues with XmlMaps
                    dynamic workbook = new Workbook(filePath);

                    // Attempt to iterate through XML maps if the collection exists
                    if (workbook.XmlMaps != null)
                    {
                        foreach (dynamic xmlMap in workbook.XmlMaps)
                        {
                            try
                            {
                                xmlMap.Refresh();
                            }
                            catch (Exception mapEx)
                            {
                                Console.Error.WriteLine($"Failed to refresh XML map in '{filePath}': {mapEx.Message}");
                            }
                        }
                    }

                    // Save the workbook back to the same file
                    workbook.Save(filePath);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }
        }

        // Creates a Windows Task Scheduler task that runs this program every night at midnight.
        private static void CreateMidnightTask()
        {
            try
            {
                const string taskName = "XmlMapNightlyRefresh";

                // Build the command line for schtasks.exe
                // /F forces creation (overwrites existing task)
                string arguments = $"/Create /F /SC DAILY /ST 00:00 /TN \"{taskName}\" /TR \"\\\"{ExecutablePath}\\\"\"";

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "schtasks",
                    Arguments = arguments,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (Process proc = Process.Start(psi))
                {
                    string output = proc.StandardOutput.ReadToEnd();
                    string error = proc.StandardError.ReadToEnd();
                    proc.WaitForExit();

                    if (proc.ExitCode != 0)
                    {
                        throw new InvalidOperationException($"schtasks failed: {error}");
                    }

                    Console.WriteLine(output);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to create scheduled task: {ex.Message}");
                throw;
            }
        }
    }
}
