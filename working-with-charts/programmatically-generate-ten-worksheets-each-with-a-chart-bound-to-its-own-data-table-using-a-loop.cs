// Title: Create ten Excel worksheets each containing a column chart with Aspose.Cells for .NET using a loop
// AI Prompts: Generate a workbook, add worksheets named Sheet1 through Sheet10, fill each with a two‑column table of sample data, and insert a column chart bound to that table using Aspose.Cells for .NET. | Write a helper method that receives a Worksheet and an index, populates cells A1:B6 with category/value rows, creates a Column chart referencing the data range, and call this method inside a for‑loop to produce charts on ten sheets.
// Common Searches: asp.net how to add a column chart to each worksheet in a loop with Aspose.Cells | c# generate multiple Excel sheets each with its own chart using Aspose.Cells | sample code for creating ten worksheets with individual charts in Aspose.Cells | loop to populate worksheets and bind charts in Aspose.Cells C# | programmatically add charts to many sheets Aspose.Cells .NET example
// Tags: Aspose.Cells create multiple worksheets with charts | C# loop add column chart per worksheet | populate worksheet data range for Aspose.Cells chart | Excel column chart generation using Aspose.Cells .NET | programmatic chart binding in Aspose.Cells workbook

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLoopDemo
{
    // // This C# program creates a new workbook, adds ten worksheets named Sheet1‑Sheet10, fills each with a simple two‑column data table, adds a column chart bound to that data, and saves the workbook as TenSheetsWithCharts.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Rename the default worksheet to "Sheet1"
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.Name = "Sheet1";

            // Populate data and add a chart to the first sheet
            PopulateSheetWithDataAndChart(firstSheet, 1);

            // Create additional 9 worksheets (total 10) and add charts
            for (int i = 2; i <= 10; i++)
            {
                // Add a new worksheet with a specific name
                Worksheet sheet = workbook.Worksheets.Add($"Sheet{i}");

                // Fill the worksheet with sample data and attach a chart
                PopulateSheetWithDataAndChart(sheet, i);
            }

            // Save the workbook to a file
            workbook.Save("TenSheetsWithCharts.xlsx");
        }

        /// <param name="sheet">Worksheet to populate.</param>
        /// <param name="sheetIndex">Index used to generate distinct data values.</param>
        private static void PopulateSheetWithDataAndChart(Worksheet sheet, int sheetIndex)
        {
            // Add header labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            // Add sample rows (5 categories)
            for (int row = 2; row <= 6; row++)
            {
                sheet.Cells[row - 1, 0].PutValue($"Item {row - 1}");
                // Generate distinct numeric values per sheet
                sheet.Cells[row - 1, 1].PutValue((row - 1) * 10 * sheetIndex);
            }

            // Add a column chart positioned within the sheet
            // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
            int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the series (values) and categories (labels)
            // Using A1-style notation; the sheet name is optional because the chart resides on the same sheet
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Optional: set a title for clarity
            chart.Title.Text = $"Chart for {sheet.Name}";
        }
    }
}
