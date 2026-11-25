// Title: Create an Excel workbook, add a worksheet, populate data, and insert a column chart using Aspose.Cells for .NET (C#)
// Common Searches: how to generate a column chart from a data range with Aspose.Cells in C# | Aspose.Cells C# example for adding a chart to a new worksheet | populate worksheet cells and bind them to a column chart using Aspose.Cells | save an Excel file with a column chart using Aspose.Cells .NET | C# code to create a workbook, add data and insert a column chart with Aspose.Cells
// Tags: Aspose.Cells add column chart to worksheet | C# populate worksheet data for chart | Aspose.Cells define series range for column chart | Aspose.Cells set chart title programmatically | Aspose.Cells save workbook as xlsx

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDemo
{
    // The sample creates a new Workbook, adds a worksheet named 'DataSheet', writes category and value data into columns A and B, adds a Column chart spanning rows 8‑21, binds the series to B2:B5 and categories to A2:A5, sets the chart title, and saves the workbook as ColumnChartDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (or add a new one)
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // Populate sample data for the chart
                // Header row
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Data rows
                string[] categories = { "A", "B", "C", "D" };
                int[] values = { 10, 25, 15, 30 };

                for (int i = 0; i < categories.Length; i++)
                {
                    // Column A: Category
                    sheet.Cells[i + 1, 0].PutValue(categories[i]); // Row index is zero‑based; i+1 points to row 2,3,...
                    // Column B: Value
                    sheet.Cells[i + 1, 1].PutValue(values[i]);
                }

                // Add a column chart to the worksheet
                // Parameters: ChartType, upper left row, upper left column, lower right row, lower right column
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Define the data range for the series (values) and categories
                // Values are in B2:B5, categories in A2:A5
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Optional: set chart title
                chart.Title.Text = "Sample Column Chart";

                // Save the workbook to a file
                string outputPath = "ColumnChartDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
