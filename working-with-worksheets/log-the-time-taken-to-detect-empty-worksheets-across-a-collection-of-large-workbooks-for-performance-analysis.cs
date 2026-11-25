// Title: Measure and log empty worksheet detection time across multiple Excel workbooks using Aspose.Cells for .NET
// Common Searches: how to log detection time for empty worksheets in a batch of Excel files using Aspose.Cells C# | performance test for empty sheet detection across large workbooks with Aspose.Cells | generate CSV report of worksheet emptiness and timing in .NET Aspose.Cells | benchmark Aspose.Cells worksheet empty check using Stopwatch in C# | process multiple workbook formats (xlsx, xlsb, csv, ods) to find empty sheets and record elapsed time
// Tags: Aspose.Cells empty worksheet detection timing | bulk Excel workbook processing with Stopwatch | worksheet emptiness CSV output Aspose.Cells | maxdatarow maxdatacolumn empty sheet check | memory optimization for large workbooks Aspose.Cells

using System;
using System.IO;
using System.Diagnostics;
using Aspose.Cells;

namespace EmptyWorksheetPerformance
{
    // The sample scans a specified folder for supported workbook files (xlsx, xls, xlsb, xlsm, csv, ods, tsv), loads each with Aspose.Cells, iterates through all worksheets, determines if a sheet is empty by checking Cells.MaxDataRow and Cells.MaxDataColumn (both -1), records the elapsed milliseconds for each worksheet using Stopwatch, writes workbook name, worksheet name, emptiness flag and detection time to a CSV file, and prints timing information to the console while disposing workbooks promptly.
    class Program
    {
        static void Main(string[] args)
        {
            // Resolve the folder that contains the workbooks
            string workbooksFolder;
            if (args.Length > 0 && Directory.Exists(args[0]))
            {
                workbooksFolder = args[0];
            }
            else
            {
                string exeDir = AppDomain.CurrentDomain.BaseDirectory;
                workbooksFolder = Path.Combine(exeDir, "Workbooks");
                if (!Directory.Exists(workbooksFolder))
                {
                    Console.Error.WriteLine($"Error: Workbooks folder not found at '{workbooksFolder}'.");
                    return;
                }
            }

            // Gather all supported workbook files (xlsx, xls, csv, ods, etc.)
            string[] supportedExtensions = new[] { ".xlsx", ".xls", ".xlsb", ".xlsm", ".csv", ".ods", ".tsv" };
            var workbookFiles = Directory.GetFiles(workbooksFolder, "*.*", SearchOption.TopDirectoryOnly);
            var filesToProcess = Array.FindAll(workbookFiles, f => Array.Exists(supportedExtensions, ext => ext.Equals(Path.GetExtension(f), StringComparison.OrdinalIgnoreCase)));

            if (filesToProcess.Length == 0)
            {
                Console.Error.WriteLine($"No workbook files found in folder '{workbooksFolder}'.");
                return;
            }

            Console.WriteLine($"Processing {filesToProcess.Length} workbook(s) in folder: {workbooksFolder}");
            Console.WriteLine();

            // Overall timer for the whole collection
            Stopwatch totalStopwatch = Stopwatch.StartNew();

            // Prepare a simple CSV report (optional)
            string reportPath = Path.Combine(workbooksFolder, "EmptyWorksheetReport.csv");
            using (var reportWriter = new StreamWriter(reportPath, false))
            {
                // CSV header
                reportWriter.WriteLine("Workbook,Worksheet,IsEmpty,DetectionTimeMs");

                foreach (string filePath in filesToProcess)
                {
                    string fileName = Path.GetFileName(filePath);
                    Console.WriteLine($"--- Processing workbook: {fileName} ---");

                    // Stopwatch for this workbook
                    Stopwatch workbookStopwatch = Stopwatch.StartNew();

                    Workbook workbook = null;
                    try
                    {
                        // Load the workbook with default load options.
                        // For very large files you may want to enable memory optimization:
                        // WorkbookSettings.MemorySetting = MemorySetting.MemoryPreference;
                        workbook = new Workbook(filePath);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Failed to load workbook '{fileName}': {ex.Message}");
                        // Record failure in report and continue with next file
                        reportWriter.WriteLine($"{fileName},,LoadFailed,0");
                        continue;
                    }

                    // Iterate through each worksheet and determine emptiness
                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                    {
                        Worksheet sheet = workbook.Worksheets[i];
                        string sheetName = sheet.Name;

                        // Use Cells.MaxDataRow and Cells.MaxDataColumn.
                        // According to the API, both return -1 when the worksheet contains no data.
                        int maxRow = sheet.Cells.MaxDataRow;
                        int maxCol = sheet.Cells.MaxDataColumn;

                        bool isEmpty = (maxRow == -1 && maxCol == -1);

                        // Time taken for this worksheet detection is negligible compared to loading,
                        // but we capture it for completeness.
                        long detectionTimeMs = workbookStopwatch.ElapsedMilliseconds;

                        // Log to console
                        Console.WriteLine($"Worksheet: '{sheetName}' | Empty: {isEmpty} | DetectionTimeMs: {detectionTimeMs}");

                        // Write to CSV
                        reportWriter.WriteLine($"{fileName},{EscapeCsv(sheetName)},{isEmpty},{detectionTimeMs}");

                        // Reset the stopwatch for the next worksheet to measure per‑worksheet time accurately
                        workbookStopwatch.Restart();
                    }

                    // Ensure workbook resources are released promptly
                    workbook.Dispose();

                    Console.WriteLine($"Finished workbook '{fileName}'. Total time for this file: {workbookStopwatch.ElapsedMilliseconds} ms");
                    Console.WriteLine();
                }

                // End of processing
                totalStopwatch.Stop();
                Console.WriteLine($"All workbooks processed. Total elapsed time: {totalStopwatch.ElapsedMilliseconds} ms");
                Console.WriteLine($"Performance report written to: {reportPath}");
            }
        }

        /// <summary>
        /// Escapes a CSV field by wrapping it in double quotes if it contains a comma, quote or newline.
        /// Double quotes inside the field are escaped by doubling them.
        /// </summary>
        private static string EscapeCsv(string field)
        {
            if (field == null)
                return "";

            bool mustQuote = field.Contains(",") || field.Contains("\"") || field.Contains("\r") || field.Contains("\n");
            if (mustQuote)
            {
                string escaped = field.Replace("\"", "\"\"");
                return $"\"{escaped}\"";
            }
            else
            {
                return field;
            }
        }
    }
}
