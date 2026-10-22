// Title: Dynamically update an Excel chart's data range based on a user‑selected date interval using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, scans column A for dates, selects rows between two DateTime variables, and assigns the matching ranges to the XValues and Values of the first chart series. | Create a C# snippet that renames a chart series to show the chosen start and end dates while updating its data source in an existing worksheet using Aspose.Cells. | Add error handling in C# to verify the input Excel file exists, ensure at least one chart is present, and gracefully handle cases where no rows match the date filter before modifying the chart with Aspose.Cells.
// Common Searches: aspnet c# filter Excel rows by date and update chart series using Aspose.Cells | change chart data source to a dynamic range based on start and end dates in Aspose.Cells | set chart series name to display selected date range programmatically with Aspose.Cells for .NET | validate chart existence before modifying series in Aspose.Cells C# example
// Tags: dynamic chart source updating Aspose.Cells | date‑range row filtering C# Aspose | configure chart series data ranges Aspose.Cells | programmatic chart series naming Aspose.Cells | ensure chart presence prior to modification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads an existing workbook, identifies rows where column A dates fall within a user‑provided start and end date, updates the first chart on the first worksheet to use those rows as the X (dates) and Y (values) ranges, renames the series to reflect the selected interval, and saves the modified workbook.
class DynamicChartUpdater
{
    static void Main()
    {
        try
        {
            // User‑selected date range
            DateTime startDate = new DateTime(2023, 1, 1);
            DateTime endDate   = new DateTime(2023, 12, 31);

            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume data is on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the rows that fall within the selected date range.
            // Data is assumed to start at row 2 (index 1) with dates in column A (index 0) and values in column B (index 1).
            int firstDataRow = 1;                     // zero‑based index of the first data row
            int lastDataRow  = cells.MaxDataRow;      // last row that contains data

            int startRow = -1;
            int endRow   = -1;

            for (int row = firstDataRow; row <= lastDataRow; row++)
            {
                // Read the date from column A
                object dateObj = cells[row, 0].Value;
                if (dateObj == null) continue;

                DateTime cellDate;
                // Try to convert the cell value to DateTime
                if (dateObj is DateTime)
                    cellDate = (DateTime)dateObj;
                else if (!DateTime.TryParse(dateObj.ToString(), out cellDate))
                    continue; // skip non‑date cells

                // Find the first row that meets the start condition
                if (startRow == -1 && cellDate >= startDate)
                    startRow = row;

                // Update the last row that still meets the end condition
                if (cellDate <= endDate)
                    endRow = row;
            }

            // If no rows match the criteria, exit
            if (startRow == -1 || endRow == -1 || endRow < startRow)
            {
                Console.WriteLine("No data found for the specified date range.");
                return;
            }

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No chart found on the worksheet.");
                return;
            }

            // Access the first chart on the worksheet (adjust index if needed)
            Chart chart = sheet.Charts[0];

            // Update each series in the chart to use the new data range
            foreach (Series series in chart.NSeries)
            {
                // Build the address strings for X (dates) and Y (values) ranges.
                // Excel addresses are 1‑based, so add 1 to row indices.
                string xRange = $"{sheet.Name}!A{startRow + 1}:A{endRow + 1}";
                string yRange = $"{sheet.Name}!B{startRow + 1}:B{endRow + 1}";

                // Set the new data source for the series
                series.XValues = xRange;
                series.Values = yRange;

                // Optionally, update the series name (e.g., based on the date range)
                series.Name = $"Data {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}";
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the updated chart
            workbook.Save(outputPath);
            Console.WriteLine("Chart data source updated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
