// Title: Insert identical column charts with individual data ranges into multiple worksheets using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, add a set of worksheets, populate each with its own A1:B6 data, and place a column chart on every sheet that references that sheet's specific range. | Change the chart type to a line chart while keeping the per‑worksheet data source logic unchanged. | Extend the loop to assign a custom chart title that includes the worksheet name and save the workbook to a filename supplied at runtime.
// Common Searches: how to add the same chart to several worksheets with different data sources using Aspose.Cells .NET | Aspose.Cells programmatically create a chart on each sheet in a workbook | set chart NSeries range per worksheet in C# Aspose.Cells | batch generate column charts for multiple sheets with unique data in Aspose.Cells
// Tags: add column chart per worksheet Aspose.Cells | set NSeries data range Aspose.Cells C# | batch chart creation Aspose.Cells .NET | dynamic chart source multiple sheets | Aspose.Cells workbook chart automation

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsBatchChartDemo
{
    // The example creates a new workbook, removes the default sheet, adds three uniquely named worksheets, fills each with sample category/value data, inserts a column chart positioned from row 8 column 1 to row 20 column 8 on each sheet, assigns the chart's data source to the sheet's own A1:B6 range, customizes the chart title with the sheet name, and saves the workbook as BatchChartsOutput.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Remove the default worksheet to avoid name collisions
                workbook.Worksheets.Clear();

                // Define the number of worksheets to process
                int sheetCount = 3;

                // Loop to add worksheets and insert identical charts with distinct data ranges
                for (int i = 0; i < sheetCount; i++)
                {
                    // Add a new worksheet with a unique name (rule: WorksheetCollection.Add(string))
                    string sheetName = $"Sheet{i + 1}";
                    Worksheet sheet = workbook.Worksheets.Add(sheetName);

                    // Populate sample data for the chart in each worksheet
                    // Header
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["B1"].PutValue("Value");

                    // Sample rows (different values per sheet to illustrate distinct ranges)
                    for (int row = 2; row <= 6; row++)
                    {
                        sheet.Cells[$"A{row}"].PutValue($"Item {row - 1}");
                        // Value varies by sheet index to make each chart unique
                        sheet.Cells[$"B{row}"].PutValue((row - 1) * (i + 1) * 10);
                    }

                    // Add a column chart to the worksheet (rule: ChartCollection.Add(ChartType, int, int, int, int))
                    // Position the chart from row 8, column 1 to row 20, column 8
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 1, 20, 8);
                    Chart chart = sheet.Charts[chartIndex];

                    // Set the data range for the chart (using NSeries.Add)
                    // The range includes the header row for proper series/category handling
                    string dataRange = $"{sheetName}!$A$1:$B$6";
                    chart.NSeries.Add(dataRange, true);

                    // Customize chart title to reflect the sheet
                    chart.Title.Text = $"Sales Data - {sheetName}";
                }

                // Save the workbook (lifecycle: save)
                workbook.Save("BatchChartsOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
