// Title: Log worksheet enumeration start and end timestamps to a file with Aspose.Cells for .NET
// Common Searches: how to write start and end timestamps to a log file while iterating worksheets using Aspose.Cells C# | Aspose.Cells performance logging example for large Excel workbooks in .NET | measure worksheet enumeration duration with Aspose.Cells and output to a text log | C# code to record enumeration time of rows in each sheet using Aspose.Cells | append timestamp and duration to a performance.log during Excel file processing with Aspose.Cells
// Tags: worksheet enumeration logging Aspose.Cells | performance monitoring Excel workbook .NET | write timestamps to log file C# | measure row iteration duration Aspose.Cells | loadoptions error handling Aspose.Cells | append performance data to text log C#

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input XLSX file, loads it with Aspose.Cells, and iterates through every worksheet up to the last data row. It creates the log directory if needed, appends a start timestamp, end timestamp, and total duration to a performance.log file, and includes robust error handling for workbook loading and I/O operations.
class PerformanceMonitor
{
    static void Main()
    {
        // Paths for the workbook and the log file
        string workbookPath = "input.xlsx";
        string logFilePath = "performance.log";

        // Verify that the workbook file exists
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Workbook file not found: {workbookPath}");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook (no password assumed; adjust LoadOptions if needed)
            LoadOptions loadOptions = new LoadOptions();
            workbook = new Workbook(workbookPath, loadOptions);
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
            return;
        }

        try
        {
            // Open the log file for appending (create directory if necessary)
            string logDir = Path.GetDirectoryName(logFilePath);
            if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
            }

            using (StreamWriter log = new StreamWriter(logFilePath, true))
            {
                // Log start timestamp
                DateTime startTime = DateTime.Now;
                log.WriteLine($"Enumeration started: {startTime:O}");

                // Enumerate all worksheets and their rows
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the last row that contains data
                    int maxDataRow = sheet.Cells.MaxDataRow;

                    // Simple enumeration: read the first cell of each row
                    for (int row = 0; row <= maxDataRow; row++)
                    {
                        Cell cell = sheet.Cells[row, 0];
                        // Optionally read the value (commented out to avoid unnecessary processing)
                        // var value = cell.Value;
                    }
                }

                // Log end timestamp and duration
                DateTime endTime = DateTime.Now;
                log.WriteLine($"Enumeration ended: {endTime:O}");
                TimeSpan duration = endTime - startTime;
                log.WriteLine($"Duration (seconds): {duration.TotalSeconds}");
                log.WriteLine(new string('-', 50));
            }
        }
        catch (IOException ex)
        {
            Console.WriteLine($"I/O error while writing log: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error during processing: {ex.Message}");
        }
    }
}
