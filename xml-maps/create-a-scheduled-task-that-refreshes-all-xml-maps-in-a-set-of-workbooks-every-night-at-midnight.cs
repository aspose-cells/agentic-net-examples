// Title: Create a Windows scheduled task in C# that nightly refreshes XML maps in Excel workbooks using Aspose.Cells
// Common Searches: C# program to schedule nightly export of Excel XML maps with Aspose.Cells | using schtasks from C# to run Aspose.Cells XML map refresh at midnight | automate XML map refresh for multiple .xlsx files in a Windows task | Aspose.Cells refresh all XmlMapCollection in a folder via scheduled task | create a daily Windows task that runs a C# console to export XML maps
// Tags: C# Windows scheduled task Aspose.Cells | export XmlMapCollection to XML using Aspose.Cells | nightly XML map refresh for Excel workbooks | process multiple .xlsx files with Aspose.Cells | schtasks integration in C# application

using System;
using System.IO;
using System.Diagnostics;
using Aspose.Cells;

// The C# console application defines a Windows scheduled task named "RefreshXmlMapsTask" that runs daily at midnight via the schtasks utility. When invoked with the "run" argument, it scans a designated "Workbooks" folder, loads each supported Excel file (.xlsx, .xls, .xlsm, .xlsb) using Aspose.Cells, iterates through its XmlMapCollection, and exports each XML map to an "ExportedXml" folder using the pattern {WorkbookName}_{XmlMapName}.xml. If the task does not exist, the code creates it automatically; otherwise it reports the existing task.
class Program
{
    // Name of the scheduled task
    private const string TaskName = "RefreshXmlMapsTask";

    // Folder that contains the workbooks to process
    private static readonly string WorkbooksFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Workbooks");

    // Folder where exported XML files will be placed
    private static readonly string ExportFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExportedXml");

    static void Main(string[] args)
    {
        try
        {
            // If the program is launched with the "run" argument, perform the refresh immediately.
            if (args.Length > 0 && args[0].Equals("run", StringComparison.OrdinalIgnoreCase))
            {
                RefreshAllXmlMaps();
                return;
            }

            // Otherwise, ensure the scheduled task exists.
            EnsureScheduledTask();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Fatal error: {ex.Message}");
        }
    }

    private static void EnsureScheduledTask()
    {
        // Determine the full path to the current executable
        string exePath = Process.GetCurrentProcess().MainModule.FileName;

        // Check if the task already exists
        bool taskExists = TaskExists(TaskName);

        if (taskExists)
        {
            Console.WriteLine($"Scheduled task \"{TaskName}\" already exists.");
            return;
        }

        // Build the schtasks command line
        // /SC DAILY   – run daily
        // /ST 00:00   – start time midnight
        // /TN         – task name
        // /TR         – command to run (quoted)
        // /F          – force creation (overwrite if exists)
        string arguments = $"/Create /SC DAILY /ST 00:00 /TN \"{TaskName}\" /TR \"\\\"{exePath}\\\" run\" /F";

        var startInfo = new ProcessStartInfo
        {
            FileName = "schtasks",
            Arguments = arguments,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using (var proc = Process.Start(startInfo))
        {
            proc.WaitForExit();

            string output = proc.StandardOutput.ReadToEnd();
            string error = proc.StandardError.ReadToEnd();

            if (proc.ExitCode == 0)
            {
                Console.WriteLine($"Scheduled task \"{TaskName}\" created successfully.");
            }
            else
            {
                Console.Error.WriteLine($"Failed to create scheduled task. ExitCode={proc.ExitCode}");
                Console.Error.WriteLine($"Output: {output}");
                Console.Error.WriteLine($"Error: {error}");
            }
        }
    }

    /// <summary>
    /// Checks whether a scheduled task with the given name exists.
    /// </summary>
    private static bool TaskExists(string taskName)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "schtasks",
            Arguments = $"/Query /TN \"{taskName}\"",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using (var proc = Process.Start(startInfo))
        {
            proc.WaitForExit();

            // schtasks returns 0 if the task is found, 1 otherwise.
            return proc.ExitCode == 0;
        }
    }

    /// <summary>
    /// Refreshes (exports) all XML maps in every workbook found in the WorkbooksFolder.
    /// The exported XML files are placed in ExportFolder with a naming convention:
    ///   {WorkbookName}_{XmlMapName}.xml
    /// </summary>
    private static void RefreshAllXmlMaps()
    {
        try
        {
            if (!Directory.Exists(WorkbooksFolder))
            {
                Console.Error.WriteLine($"Workbooks folder does not exist: {WorkbooksFolder}");
                return;
            }

            Directory.CreateDirectory(ExportFolder);

            string[] workbookFiles = Directory.GetFiles(WorkbooksFolder, "*.*", SearchOption.TopDirectoryOnly);
            if (workbookFiles.Length == 0)
            {
                Console.WriteLine("No workbook files found to process.");
                return;
            }

            foreach (string filePath in workbookFiles)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                // Process only Excel formats supported by Aspose.Cells
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm" && extension != ".xlsb")
                {
                    Console.WriteLine($"Skipping unsupported file: {Path.GetFileName(filePath)}");
                    continue;
                }

                ProcessWorkbook(filePath);
            }

            Console.WriteLine("All XML maps have been refreshed.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error during refresh operation: {ex.Message}");
        }
    }

    /// <summary>
    /// Loads a workbook, iterates through its XML maps, and exports each map to an XML file.
    /// </summary>
    private static void ProcessWorkbook(string workbookPath)
    {
        try
        {
            Console.WriteLine($"Processing workbook: {Path.GetFileName(workbookPath)}");

            // Load the workbook
            Workbook wb = new Workbook(workbookPath);

            // Access the collection of XML maps for the workbook (they are stored at the worksheet collection level)
            XmlMapCollection xmlMaps = wb.Worksheets.XmlMaps;

            if (xmlMaps == null || xmlMaps.Count == 0)
            {
                Console.WriteLine("  No XML maps found in this workbook.");
                return;
            }

            for (int i = 0; i < xmlMaps.Count; i++)
            {
                XmlMap map = xmlMaps[i];
                string mapName = map.Name;

                // Build output file name
                string safeMapName = MakeFileSystemSafe(mapName);
                string outputFileName = $"{Path.GetFileNameWithoutExtension(workbookPath)}_{safeMapName}.xml";
                string outputPath = Path.Combine(ExportFolder, outputFileName);

                // Export the XML data linked to this map
                wb.ExportXml(mapName, outputPath);
                Console.WriteLine($"  Exported XML map \"{mapName}\" to \"{outputFileName}\"");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"  Failed to process workbook \"{Path.GetFileName(workbookPath)}\": {ex.Message}");
        }
    }

    /// <summary>
    /// Replaces characters that are invalid in file names.
    /// </summary>
    private static string MakeFileSystemSafe(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
