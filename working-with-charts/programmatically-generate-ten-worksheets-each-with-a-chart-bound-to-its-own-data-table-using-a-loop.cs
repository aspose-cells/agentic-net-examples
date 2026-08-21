// Title: Create 10 Worksheets with Individual Column Charts Using Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to generate a workbook, remove the default sheet, and use a for‑loop to add ten worksheets (Sheet1‑Sheet10). Each sheet receives a header, five rows of sample data, and a column chart positioned from row 8 column 0 to row 25 column 8. The chart series are bound to B2:B6, categories to A2:A6, and a title such as "Chart for Sheet3" is applied. The workbook is saved as TenSheetsWithCharts.xlsx.
// Keywords: Aspose.Cells C# example | create multiple worksheets Aspose | add chart in loop .NET | column chart programmatically | Excel automation Aspose.Cells | generate charts per sheet | Aspose.Cells tutorial | GitHub Aspose.Cells chart loop | C# Excel chart binding | Aspose.Cells workbook with charts
// Common Searches: how to add a chart to each worksheet using Aspose.Cells C# | Aspose.Cells loop to create multiple sheets with charts | C# generate Excel file with ten sheets and individual charts | programmatically bind chart series and categories in Aspose.Cells | example of column chart per worksheet Aspose.Cells
// Developer Intent: Produce an Excel workbook containing ten worksheets, each with its own column chart linked to sheet‑specific data.
// Use Cases: Monthly sales report where each month occupies a separate sheet with a dedicated chart. | Multi‑sheet dashboard that visualizes distinct data sets on individual worksheets. | Automated batch generation of Excel files where every data segment requires its own chart.
// AI Prompts: Generate C# code with Aspose.Cells that creates N worksheets, each containing a line chart bound to unique data ranges. | Adapt the loop to assign a different chart type (e.g., pie, bar, scatter) to each worksheet based on its index. | Enhance the example by adding data labels, a legend, and custom colors to every chart inside the loop, then save the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLoopDemo
{
    // This C# example demonstrates how to generate a workbook, remove the default sheet, and use a for‑loop to add ten worksheets (Sheet1‑Sheet10). Each sheet receives a header, five rows of sample data, and a column chart positioned from row 8 column 0 to row 25 column 8. The chart series are bound to B2:B6, categories to A2:A6, and a title such as "Chart for Sheet3" is applied. The workbook is saved as TenSheetsWithCharts.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (default contains one worksheet)
                Workbook workbook = new Workbook();

                // Remove the default worksheet to avoid name conflicts
                workbook.Worksheets.Clear();

                // Loop to create 10 worksheets, each with its own data and chart
                for (int i = 1; i <= 10; i++)
                {
                    // Add a new worksheet with a distinct name
                    string sheetName = $"Sheet{i}";
                    Worksheet sheet = workbook.Worksheets.Add(sheetName);

                    // Populate sample data for the chart
                    // Header row
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["B1"].PutValue("Value");

                    // Data rows (5 rows of sample data)
                    for (int row = 2; row <= 6; row++)
                    {
                        sheet.Cells[row - 1, 0].PutValue($"Item {row - 1}");
                        // Example value varies per sheet to make each chart unique
                        sheet.Cells[row - 1, 1].PutValue((row - 1) * 10 + i);
                    }

                    // Add a column chart to the worksheet
                    // Position the chart from row 8, column 0 to row 25, column 8
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 8);
                    Chart chart = sheet.Charts[chartIndex];

                    // Set the data range for the series (values) and categories
                    // Values are in B2:B6, categories in A2:A6
                    chart.NSeries.Add("B2:B6", true);
                    chart.NSeries.CategoryData = "A2:A6";

                    // Optional: give the chart a title
                    chart.Title.Text = $"Chart for {sheetName}";
                }

                // Save the workbook to a file
                workbook.Save("TenSheetsWithCharts.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
