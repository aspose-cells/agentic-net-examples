using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the large workbook to be processed
        string workbookPath = "LargeWorkbook.xlsx";

        // Path to the log file where timestamps will be recorded
        string logPath = "PerformanceLog.txt";

        // Open the log file (overwrite if it exists)
        using (StreamWriter log = new StreamWriter(logPath, false))
        {
            // Load the workbook (no special load options required for this demo)
            Workbook workbook = new Workbook(workbookPath);

            // Log the overall start time of the enumeration
            log.WriteLine($"Workbook enumeration started at {DateTime.Now:O}");

            // Iterate through each worksheet in the workbook
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIndex];

                // Log the start time for the current worksheet
                log.WriteLine($"Worksheet '{sheet.Name}' (index {sheetIndex}) start at {DateTime.Now:O}");

                // Example enumeration: iterate over rows that contain data
                Cells cells = sheet.Cells;
                int maxDataRow = cells.MaxDataRow; // last row with data

                for (int row = 0; row <= maxDataRow; row++)
                {
                    // Access a cell to ensure the row is actually read.
                    // In a real scenario you would perform your processing here.
                    var _ = cells[row, 0].Value;
                }

                // Log the end time for the current worksheet
                log.WriteLine($"Worksheet '{sheet.Name}' (index {sheetIndex}) end at {DateTime.Now:O}");
            }

            // Log the overall end time of the enumeration
            log.WriteLine($"Workbook enumeration finished at {DateTime.Now:O}");
        }
    }
}