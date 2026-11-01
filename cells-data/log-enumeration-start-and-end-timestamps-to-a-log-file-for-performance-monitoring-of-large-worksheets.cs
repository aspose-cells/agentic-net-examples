// Title: Append UTC start and end timestamps to a log file while enumerating used cells of a large worksheet with Aspose.Cells in C#
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, iterates over the used range of the first worksheet, records the UTC start and end times, calculates the elapsed seconds, and appends these values to a text log file. | Adjust the enumeration loop so that after logging the processing duration a separator line is written, ensuring the log file remains in append mode.
// Common Searches: C# Aspose.Cells how to log enumeration start and end time for large Excel worksheet | measure performance of worksheet cell iteration using Aspose.Cells .NET | append processing timestamps to a text file while reading Excel cells with Aspose.Cells | log total duration of used range traversal in Aspose.Cells C# example | performance monitoring large workbook enumeration Aspose.Cells logging timestamps
// Tags: Aspose.Cells worksheet enumeration timing | log timestamps during Excel cell traversal .NET | measure used range iteration performance Aspose.Cells | append UTC timestamps to text log C# | performance monitoring large workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads a large Excel workbook with Aspose.Cells, iterates over the used cells of the first worksheet, and appends UTC start time, end time, and total duration to a performance log file.
class WorksheetPerformanceLogger
{
    static void Main()
    {
        // Path to the Excel file to be processed
        string excelPath = "LargeWorkbook.xlsx";

        // Path to the log file where timestamps will be recorded
        string logPath = "PerformanceLog.txt";

        // Open the log file for appending timestamps
        using (StreamWriter logWriter = new StreamWriter(logPath, true))
        {
            // Log the start timestamp
            DateTime startTime = DateTime.UtcNow;
            logWriter.WriteLine($"Enumeration Start: {startTime:O}");

            // Load the workbook using Aspose.Cells (lifecycle rule: load)
            Workbook workbook = new Workbook(excelPath);

            // Assume we are interested in the first worksheet (adjust as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the used range to limit enumeration to actual data
            int maxRow = worksheet.Cells.MaxDataRow;
            int maxColumn = worksheet.Cells.MaxDataColumn;

            // Enumerate through rows and columns (example of large worksheet processing)
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    // Access each cell (placeholder for actual processing logic)
                    Cell cell = worksheet.Cells[row, col];
                    // Example: read the cell value (no operation performed)
                    var value = cell.Value;
                }
            }

            // Log the end timestamp
            DateTime endTime = DateTime.UtcNow;
            logWriter.WriteLine($"Enumeration End: {endTime:O}");

            // Optionally, log the elapsed time
            TimeSpan duration = endTime - startTime;
            logWriter.WriteLine($"Total Duration: {duration.TotalSeconds:F2} seconds");
            logWriter.WriteLine(new string('-', 50));
        }

        // If any modifications were made and need to be saved, use the save lifecycle rule
        // workbook.Save("ProcessedWorkbook.xlsx");
    }
}
