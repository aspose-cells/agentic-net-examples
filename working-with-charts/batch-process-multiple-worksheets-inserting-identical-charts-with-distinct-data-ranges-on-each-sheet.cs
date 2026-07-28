// Title: Add identical column charts with sheet‑specific data to multiple worksheets using Aspose.Cells for .NET
// Description: Shows how to create a workbook, fill several worksheets with a header and sample values, and insert a column chart on each sheet that reads the sheet’s own A1:B5 range. The chart is placed at rows 7‑20, columns A‑I, and its title is set to the worksheet name before saving.
// Keywords: Aspose.Cells | C# chart batch | add chart to multiple worksheets | Excel column chart Aspose | loop create charts .NET | populate worksheets with data Aspose.Cells | batch chart generation | Excel automation C# | Aspose.Cells chart example
// Common Searches: Aspose.Cells add same chart to each worksheet | C# loop to create charts on multiple Excel sheets | set chart data source per sheet Aspose.Cells | batch generate Excel charts .NET | how to copy chart to all worksheets Aspose | Excel chart automation C# Aspose.Cells
// Developer Intent: Create the same chart type on every worksheet while each chart uses the data from its own sheet.
// Use Cases: Regional sales workbook where each sheet represents a region and contains a column chart of that region’s sales. | Monthly financial dashboard that adds a uniform chart layout to each month’s worksheet, pulling that month’s figures from the same cell range. | Product‑category report where every category sheet includes a chart visualizing its own data set using a shared template.
// AI Prompts: Generate C# code with Aspose.Cells that adds a line chart to every worksheet in an existing workbook, using range A2:C10 and setting the chart title to the sheet name. | Refactor the batch‑chart example into a reusable method that accepts chart type, data range, and position parameters. | Explain how to offset chart positions on successive worksheets while keeping the same data source range in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsBatchChartDemo
{
    // Shows how to create a workbook, fill several worksheets with a header and sample values, and insert a column chart on each sheet that reads the sheet’s own A1:B5 range. The chart is placed at rows 7‑20, columns A‑I, and its title is set to the worksheet name before saving.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Number of worksheets to process
            int sheetCount = 3;

            // Populate each worksheet with sample data and add a chart
            for (int i = 0; i < sheetCount; i++)
            {
                // Add a new worksheet (or use the first default sheet for i == 0)
                Worksheet sheet = i == 0 ? workbook.Worksheets[0] : workbook.Worksheets.Add($"Sheet{i + 1}");

                // Define data range for this sheet (e.g., A1:B5, offset by sheet index)
                // Here we simply fill the same range with different values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                for (int row = 2; row <= 5; row++)
                {
                    sheet.Cells[$"A{row}"].PutValue($"Item{row - 1}");
                    // Vary the numeric value per sheet to demonstrate distinct data
                    sheet.Cells[$"B{row}"].PutValue((row - 1) * 10 + i * 5);
                }

                // Add a column chart to the worksheet at a fixed position
                // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart specific to this sheet
                // The range includes the header row and data rows
                string dataRange = $"A1:B5";
                chart.NSeries.Add(dataRange, true);

                // Optional: customize chart title to reflect the sheet name
                chart.Title.Text = $"Sales Data - {sheet.Name}";
            }

            // Save the workbook with all charts inserted
            workbook.Save("BatchChartsOutput.xlsx");
        }
    }
}
