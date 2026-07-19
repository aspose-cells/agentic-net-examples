// Title: Hide hidden rows from an Aspose.Cells chart using PlotVisibleCellsOnly (C#)
// Description: Demonstrates how to create a workbook, hide specific rows, add a column chart, and set the PlotVisibleCellsOnly property so that hidden rows are omitted from the chart series while the original data range remains unchanged.
// Keywords: Aspose.Cells PlotVisibleCellsOnly | C# hide rows chart | exclude hidden data points Aspose.Cells | chart series visibility hidden rows | Aspose.Cells column chart hide categories
// Common Searches: Aspose.Cells hide rows in chart | PlotVisibleCellsOnly C# example | remove hidden data points from chart Aspose.Cells | chart series visibility false hidden rows | how to exclude hidden rows from Aspose.Cells chart
// Developer Intent: Exclude data points that belong to hidden worksheet rows from a chart without modifying the source range.
// Use Cases: Financial reports where future periods are hidden in the sheet and the chart automatically skips those columns. | Interactive dashboards that let users hide rows to filter chart data while preserving the underlying dataset. | Printable reports where categories marked as hidden are removed from the visual chart but stay in the data source for calculations.
// AI Prompts: Generate C# code that toggles PlotVisibleCellsOnly at runtime to show or hide hidden data points in an existing Aspose.Cells chart. | Show how to hide specific data points in a pie chart based on hidden worksheet rows using Aspose.Cells for .NET. | Explain how PlotVisibleCellsOnly affects different chart types when rows are hidden in the worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, hide specific rows, add a column chart, and set the PlotVisibleCellsOnly property so that hidden rows are omitted from the chart series while the original data range remains unchanged.
    public class HideHiddenDataPointsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data (categories in column A, values in column B)
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B4"].PutValue(30);
                worksheet.Cells["A5"].PutValue("D");
                worksheet.Cells["B5"].PutValue(40);

                // Hide rows that contain data points we want to make invisible (e.g., rows 3 and 5)
                worksheet.Cells.Rows[2].IsHidden = true; // hides row 3 (category B)
                worksheet.Cells.Rows[4].IsHidden = true; // hides row 5 (category D)

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B5", true);          // values
                chart.NSeries.CategoryData = "A2:A5";     // categories

                // Plot only visible cells
                chart.PlotVisibleCellsOnly = true;

                // Save the workbook
                string outputPath = "HideHiddenDataPointsDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HideHiddenDataPointsDemo.Run();
        }
    }
}
