// Title: Log enumeration start and end timestamps for each worksheet in a large Excel workbook using Aspose.Cells for .NET
// AI Prompts: Create C# code that opens an .xlsx file with Aspose.Cells, loops through all worksheets, and writes each sheet's enumeration start and finish times to a log file. | Modify the enumeration loop to capture the total row and column count processed per worksheet and append these metrics to the same log. | Add logic to record the overall workbook processing duration and include it at the end of the log file.
// Common Searches: aspnet log worksheet enumeration time Aspose.Cells large workbook | how to measure sheet processing duration with Aspose.Cells in C# | record start and end timestamps for each Excel sheet using Aspose.Cells .NET | performance logging of cell enumeration across worksheets Aspose.Cells
// Tags: Aspose.Cells worksheet enumeration timing log | log sheet processing timestamps C# | measure Excel worksheet performance Aspose | record worksheet enumeration duration .NET | large workbook cell enumeration monitoring

using System;
using System.IO;
using Aspose.Cells;

// The example loads a large Excel workbook with Aspose.Cells, iterates through each worksheet, and writes the start time, end time, and duration of cell enumeration for every sheet to a text file, enabling performance monitoring of worksheet processing.
class WorksheetEnumerationLogger
{
    static void Main()
    {
        // Path to the large workbook to be processed
        string workbookPath = "LargeWorkbook.xlsx";

        // Path to the log file where timestamps will be recorded
        string logPath = "EnumerationLog.txt";

        // Initialize the log file (overwrite if it already exists)
        using (var initLog = new StreamWriter(logPath, false))
        {
            initLog.WriteLine($"Enumeration started at {DateTime.Now:O}");
        }

        // Load the workbook (standard loading; can be replaced with LoadOptions for large files)
        Workbook workbook = new Workbook(workbookPath);

        // Open the log file for appending timestamps during enumeration
        using (var logWriter = new StreamWriter(logPath, true))
        {
            // Iterate through each worksheet in the workbook
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIndex];

                // Record start time for the current worksheet
                DateTime sheetStart = DateTime.Now;
                logWriter.WriteLine($"Sheet {sheetIndex} ('{sheet.Name}') enumeration start: {sheetStart:O}");

                // Access the Cells collection to enumerate rows and columns
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;       // Highest row index containing data
                int maxCol = cells.MaxDataColumn;    // Highest column index containing data

                // Simple enumeration of all cells that have data
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Access the cell; reading its value avoids heavy processing
                        Cell cell = cells[row, col];
                        var value = cell.Value;
                    }
                }

                // Record end time for the current worksheet
                DateTime sheetEnd = DateTime.Now;
                logWriter.WriteLine($"Sheet {sheetIndex} ('{sheet.Name}') enumeration end: {sheetEnd:O}");
                logWriter.WriteLine($"Sheet {sheetIndex} duration (ms): {(sheetEnd - sheetStart).TotalMilliseconds}");
            }

            // Record overall completion time
            DateTime overallEnd = DateTime.Now;
            logWriter.WriteLine($"Enumeration finished at {overallEnd:O}");
        }
    }
}
