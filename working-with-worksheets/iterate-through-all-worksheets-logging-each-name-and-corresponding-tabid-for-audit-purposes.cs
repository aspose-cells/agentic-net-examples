// Title: How to audit all worksheets in an Excel workbook by logging each sheet name and its TabId using Aspose.Cells for .NET
// Common Searches: C# code to list worksheet names with TabId using Aspose.Cells | Aspose.Cells .NET iterate worksheets and write TabId to a log file | How to get TabId of each sheet in an Excel workbook with Aspose.Cells | Create audit log of worksheet identifiers (name and TabId) in C# | Retrieve worksheet TabId property in Aspose.Cells example
// Tags: Aspose.Cells worksheet TabId enumeration | C# audit log for worksheet names | Excel workbook worksheet iteration Aspose.Cells | retrieve worksheet identifier Aspose.Cells .NET | log worksheet TabId to file using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetTabIdAudit
{
    // The sample loads a specified Excel workbook with Aspose.Cells, iterates through every worksheet, and records each worksheet's Name and TabId to both the console and an audit log file, handling missing files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the workbook file path
            string workbookPath;
            if (args.Length > 0 && !string.IsNullOrWhiteSpace(args[0]))
            {
                workbookPath = args[0];
            }
            else
            {
                workbookPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sample.xlsx");
                Console.WriteLine($"No workbook path supplied. Trying default location: {workbookPath}");
            }

            // Validate the file existence
            if (!File.Exists(workbookPath))
            {
                Console.Error.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Prepare audit log file (overwrites any existing file)
                string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WorksheetTabIdAuditLog.txt");
                using (StreamWriter logWriter = new StreamWriter(logFilePath, false))
                {
                    logWriter.WriteLine($"Audit Log - {DateTime.UtcNow:u}");
                    logWriter.WriteLine($"Workbook: {Path.GetFileName(workbookPath)}");
                    logWriter.WriteLine(new string('-', 50));

                    // Iterate through all worksheets
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        string line = $"Worksheet Name: \"{sheet.Name}\", TabId: {sheet.TabId}";
                        Console.WriteLine(line);
                        logWriter.WriteLine(line);
                    }

                    logWriter.WriteLine(new string('-', 50));
                    logWriter.WriteLine("Audit completed successfully.");
                }

                Console.WriteLine($"Audit log written to: {logFilePath}");
            }
            catch (Exception ex)
            {
                // Capture any unexpected errors
                Console.Error.WriteLine($"An exception occurred while processing the workbook: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
