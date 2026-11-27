// Title: C# example: Log original and new Worksheet TabId values with Aspose.Cells and thread‑safe file output
// Common Searches: Aspose.Cells C# how to log worksheet TabId changes to a file | record original and new TabId when updating Excel sheet tabs using Aspose.Cells | thread‑safe logging of worksheet TabId property changes in .NET | sample code for setting and auditing Worksheet TabId with Aspose.Cells | C# program to change Excel worksheet TabId and write audit log
// Tags: Aspose.Cells worksheet TabId audit logging | C# thread‑safe file logging for Excel sheet identifiers | set worksheet TabId with Aspose.Cells | log Excel worksheet TabId changes | Aspose.Cells workbook TabId modification example

using System;
using System.IO;
using Aspose.Cells;

namespace TabIdLoggingDemo
{
    // The TabIdLogger class demonstrates a thread‑safe way to change a worksheet's TabId using Aspose.Cells for .NET while writing an audit entry that includes the worksheet name, original TabId, new TabId, and UTC timestamp to the console and optionally to a specified log file. The accompanying program loads or creates a workbook, applies deterministic TabId updates to each sheet, saves the workbook, and verifies the persisted values.
    internal static class TabIdLogger
    {
        private static readonly object _lock = new object();

        /// <summary>
        /// Logs a TabId change.
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet whose TabId is being changed.</param>
        /// <param name="originalTabId">The original TabId value.</param>
        /// <param name="newTabId">The new TabId value.</param>
        /// <param name="logFilePath">Optional full path to a log file. If null or empty, file logging is skipped.</param>
        public static void LogChange(string worksheetName, int originalTabId, int newTabId, string logFilePath = null)
        {
            string message = $"[{DateTime.UtcNow:O}] Worksheet \"{worksheetName}\": TabId changed from {originalTabId} to {newTabId}";
            // Write to console
            Console.WriteLine(message);

            // Write to file if a path is supplied
            if (!string.IsNullOrWhiteSpace(logFilePath))
            {
                try
                {
                    lock (_lock)
                    {
                        File.AppendAllText(logFilePath, message + Environment.NewLine);
                    }
                }
                catch (Exception ex)
                {
                    // Do not hide the failure – report it to the console.
                    Console.Error.WriteLine($"Failed to write log entry to \"{logFilePath}\": {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Sets a new TabId on the supplied worksheet and logs the change.
        /// </summary>
        /// <param name="worksheet">The worksheet whose TabId will be changed.</param>
        /// <param name="newTabId">The new TabId value to assign.</param>
        /// <param name="logFilePath">Optional log file path.</param>
        public static void SetTabId(Worksheet worksheet, int newTabId, string logFilePath = null)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));

            int originalTabId = worksheet.TabId;
            if (originalTabId == newTabId)
            {
                // No change – still log for completeness.
                LogChange(worksheet.Name, originalTabId, newTabId, logFilePath);
                return;
            }

            worksheet.TabId = newTabId;
            LogChange(worksheet.Name, originalTabId, newTabId, logFilePath);
        }
    }

    internal class Program
    {
        /// <summary>
        /// Entry point.
        /// Usage:
        ///   TabIdLoggingDemo.exe <inputExcelPath> <outputExcelPath> [logFilePath]
        /// If <inputExcelPath> does not exist, a new workbook with a single sheet is created.
        /// </summary>
        private static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.Error.WriteLine("Insufficient arguments.");
                Console.Error.WriteLine("Usage: TabIdLoggingDemo.exe <inputExcelPath> <outputExcelPath> [logFilePath]");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];
            string logPath = args.Length >= 3 ? args[2] : null;

            // Ensure the directory for the output file exists
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to prepare output directory: {ex.Message}");
                return;
            }

            // Load or create workbook
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Failed to load workbook \"{inputPath}\": {ex.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook with a default sheet.");
                workbook = new Workbook(); // creates a workbook with one default worksheet
            }

            // Demonstrate changing TabId values.
            // For this demo we will assign a deterministic new TabId:
            //   newTabId = originalTabId + 1000 * (worksheet index + 1)
            // This ensures each sheet gets a unique new TabId.
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet ws = workbook.Worksheets[i];
                int originalTabId = ws.TabId;
                int newTabId = originalTabId + 1000 * (i + 1);
                try
                {
                    TabIdLogger.SetTabId(ws, newTabId, logPath);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error setting TabId for worksheet \"{ws.Name}\": {ex.Message}");
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to save workbook to \"{outputPath}\": {ex.Message}");
                return;
            }

            // Verify that the TabId values were persisted
            try
            {
                Workbook verificationWb = new Workbook(outputPath);
                Console.WriteLine("Verification of saved TabId values:");
                foreach (Worksheet ws in verificationWb.Worksheets)
                {
                    Console.WriteLine($"Worksheet \"{ws.Name}\" – TabId = {ws.TabId}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to verify saved workbook: {ex.Message}");
            }
        }
    }
}
