// Title: Identify and log Excel worksheets that contain only header rows using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that opens an .xlsx workbook, scans each worksheet, detects sheets where only the first row has data, and writes the sheet names with a timestamp to a text log file. | Create a .NET method that returns a list of worksheet names that have no data beyond the header row and saves a review report to a specified path.
// Common Searches: Aspose.Cells C# find worksheets with only header row | how to generate a log of Excel sheets that contain only column headings using Aspose.Cells | detect Excel worksheets that have no data rows beyond the first row in .NET | flag header‑only worksheets in an .xlsx file and export a report | C# scan workbook for sheets with only header data and write to text file
// Tags: Aspose.Cells detect header‑only worksheets | C# scan workbook for empty data rows | generate review log for Excel sheets | log header‑only worksheets .NET | Worksheet MaxDataRow usage Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorksheetHeaderChecker
{
    // The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet, uses MaxDataRow/MaxDataColumn to determine if only the first row contains data, collects those sheet names, and writes a timestamped report to both the console and a text file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path
            string inputPath = @"C:\Data\InputWorkbook.xlsx";

            // Output log file path
            string logPath = @"C:\Data\HeaderReviewLog.txt";

            try
            {
                // Verify input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(inputPath);

                // List to hold worksheets that contain only header rows
                List<string> headerOnlySheets = new List<string>();

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the last row and column that contain data (zero‑based indices)
                    int maxDataRow = sheet.Cells.MaxDataRow;
                    int maxDataColumn = sheet.Cells.MaxDataColumn;

                    // Skip completely empty worksheets
                    if (maxDataRow < 0 || maxDataColumn < 0)
                        continue;

                    // If the only populated row is the first row (index 0), it's a header‑only sheet
                    if (maxDataRow == 0)
                    {
                        // Verify that the first row actually contains at least one non‑empty cell
                        bool hasHeaderData = false;
                        for (int col = 0; col <= maxDataColumn; col++)
                        {
                            Cell cell = sheet.Cells[0, col];
                            if (cell.Value != null && !string.IsNullOrEmpty(cell.StringValue))
                            {
                                hasHeaderData = true;
                                break;
                            }
                        }

                        if (hasHeaderData)
                            headerOnlySheets.Add(sheet.Name);
                    }
                }

                // Prepare log details
                List<string> logLines = new List<string>
                {
                    $"Report generated on {DateTime.Now}",
                    $"Total worksheets scanned: {workbook.Worksheets.Count}",
                    $"Worksheets flagged for header‑only content: {headerOnlySheets.Count}",
                    string.Empty
                };

                foreach (string sheetName in headerOnlySheets)
                {
                    logLines.Add($"- Worksheet \"{sheetName}\" contains only header rows and requires review.");
                }

                // Write log to console
                foreach (string line in logLines)
                {
                    Console.WriteLine(line);
                }

                // Ensure the directory for the log file exists
                string logDir = Path.GetDirectoryName(logPath);
                if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }

                // Write log to file
                File.WriteAllLines(logPath, logLines);
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
