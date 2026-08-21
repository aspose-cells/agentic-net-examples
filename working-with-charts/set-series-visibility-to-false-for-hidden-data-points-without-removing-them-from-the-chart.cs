// Title: Aspose.Cells C# – Hide Hidden Rows in a Chart While Keeping the Series Intact
// Description: Demonstrates how to create a workbook, hide specific rows, add a column chart, and use the PlotVisibleCellsOnly property so the chart displays only visible data points without removing them from the series. The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells PlotVisibleCellsOnly | C# hide hidden rows chart | exclude hidden data points Aspose.Cells | chart visible cells only C# | Aspose.Cells series visibility | Aspose.Cells hide rows without deleting series | Aspose.Cells chart filtering hidden rows
// Common Searches: Aspose.Cells hide hidden rows from chart | C# PlotVisibleCellsOnly example | how to exclude hidden rows in Aspose.Cells chart | Aspose.Cells chart show only visible cells | remove hidden data points from chart C#
// Developer Intent: Show hidden rows in a worksheet but prevent their values from appearing in a chart, while preserving the original series definition.
// Use Cases: Financial reports where confidential rows are hidden but must not be plotted. | Sales dashboards that automatically ignore rows filtered out by the user. | Inventory sheets where discontinued items are hidden and should not affect chart trends.
// AI Prompts: Generate C# code using Aspose.Cells to create a line chart that plots only visible cells, ignoring hidden rows. | Explain how the PlotVisibleCellsOnly property impacts different chart types in Aspose.Cells and how to enable or disable it. | Provide a sample that hides multiple rows and updates a pie chart so hidden slices are omitted from the visual output.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, hide specific rows, add a column chart, and use the PlotVisibleCellsOnly property so the chart displays only visible data points without removing them from the series. The workbook is saved as an XLSX file.
    public class HideHiddenDataPointsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (categories in column A, values in column B)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Hide the row that contains the second data point (row 3 -> index 2)
                sheet.Cells.Rows[2].IsHidden = true;

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIdx];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Ensure that only visible cells are plotted.
                chart.PlotVisibleCellsOnly = true; // default is true, set explicitly for clarity

                // Save the workbook
                string outputPath = "HideHiddenDataPointsDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
