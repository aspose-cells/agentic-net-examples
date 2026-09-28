// Title: Add a clustered column chart to every worksheet in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# that loops through all worksheets in a Workbook, inserts a clustered column chart using the range A1:B5, and saves the file as a new workbook with Aspose.Cells. | Demonstrate how to assign each chart’s title to the corresponding worksheet name and set a series name using the Aspose.Cells chart API.
// Common Searches: asp.net create column chart on each sheet using Aspose.Cells | c# loop through worksheets and add the same chart Aspose.Cells example | how to set chart title to worksheet name in Aspose.Cells .NET | generate clustered column charts for multiple worksheets programmatically | Aspose.Cells add chart from range A1:B5 to every worksheet
// Tags: add column chart per worksheet Aspose.Cells | loop worksheets generate charts .NET | clustered column chart range A1:B5 Aspose.Cells | set chart title to worksheet name Aspose.Cells | save workbook with charts Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a sample workbook if missing, loads it, iterates over each worksheet, adds a clustered column chart that references cells A1:B5, sets the series name, customizes the chart title with the sheet name, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Name = "Sheet1";

                // Sample data in A1:B5
                for (int i = 0; i < 5; i++)
                {
                    ws.Cells[i, 0].PutValue($"Category {i + 1}");
                    ws.Cells[i, 1].PutValue(i + 1);
                }

                wb.Save(inputPath);
            }

            // Load the workbook.
            var workbook = new Workbook(inputPath);

            // Iterate over each worksheet and add a column chart.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add a clustered column chart at the specified position.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Define the data range (A1:B5).
                int firstRow = 0;      // zero‑based row index
                int firstColumn = 0;   // zero‑based column index (A)
                int totalRows = 5;     // number of rows in the range

                string categoryRange = $"{sheet.Name}!${CellsHelper.ColumnIndexToName(firstColumn)}${firstRow + 1}:${CellsHelper.ColumnIndexToName(firstColumn)}${firstRow + totalRows}";
                string valuesRange = $"{sheet.Name}!${CellsHelper.ColumnIndexToName(firstColumn + 1)}${firstRow + 1}:${CellsHelper.ColumnIndexToName(firstColumn + 1)}${firstRow + totalRows}";

                // Add the series to the chart (valuesRange, true indicates data is in rows/columns as needed).
                int seriesIndex = chart.NSeries.Add(valuesRange, true);
                // Set series name.
                chart.NSeries[seriesIndex].Name = "Series";

                // NOTE: CategoryData property may not be available in some versions.
                // If needed, uncomment the following line after confirming the API supports it.
                // chart.NSeries[seriesIndex].CategoryData = categoryRange;

                // Set a title for the chart.
                chart.Title.Text = $"Chart for {sheet.Name}";
            }

            // Save the workbook with the newly created charts.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
