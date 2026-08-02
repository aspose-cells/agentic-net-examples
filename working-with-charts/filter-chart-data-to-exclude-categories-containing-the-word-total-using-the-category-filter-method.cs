// Title: Aspose.Cells .NET: Hide “Total” categories in a chart by filtering rows and using PlotVisibleCellsOnly
// Description: Creates a workbook, fills it with category/value data, adds a column chart, hides rows whose category contains the word “Total”, sets PlotVisibleCellsOnly so the chart displays only visible categories, and saves the file as ChartFilteredByCategory.xlsx.
// Keywords: Aspose.Cells | C# chart filtering | Hide rows by category | PlotVisibleCellsOnly | Exclude Total rows | Category filter method | Excel chart data filter | .NET workbook chart | Aspose.Cells example | ChartCategoryFilter
// Common Searches: Aspose.Cells hide rows containing 'Total' in chart | PlotVisibleCellsOnly chart Aspose.Cells .NET | Exclude Total categories from Excel chart using Aspose | Filter chart data by category text Aspose.Cells | C# code to remove subtotal rows from chart
// Developer Intent: Programmatically hide rows whose category label includes the word “Total” and configure the chart to render only the visible categories using Aspose.Cells for .NET.
// Use Cases: Sales region report where subtotal “Total” rows should not appear in the column chart. | Financial statement where aggregated totals are hidden to highlight individual line items. | KPI dashboard that automatically omits total rows when generating Excel charts. | Automated reporting pipeline that filters out summary rows before chart creation.
// AI Prompts: Give a step‑by‑step guide to hide rows with a specific keyword and enable PlotVisibleCellsOnly in Aspose.Cells C#. | Show how to extend the example to filter multiple keywords such as 'Total', 'Subtotal', or 'Grand Total'. | Explain how to temporarily unhide the hidden rows after saving the workbook while keeping the chart unchanged. | Provide a GitHub‑style snippet that uses LINQ to identify rows to hide based on category text.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCategoryFilter
{
    // Creates a workbook, fills it with category/value data, adds a column chart, hides rows whose category contains the word “Total”, sets PlotVisibleCellsOnly so the chart displays only visible categories, and saves the file as ChartFilteredByCategory.xlsx.
    public class FilterChartCategories
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with some categories containing the word "Total"
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("North");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("South");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("East Total");
                sheet.Cells["B4"].PutValue(200);
                sheet.Cells["A5"].PutValue("West");
                sheet.Cells["B5"].PutValue(130);
                sheet.Cells["A6"].PutValue("Central Total");
                sheet.Cells["B6"].PutValue(180);

                // Add a column chart based on the data range
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B6", true);          // Values
                chart.NSeries.CategoryData = "A2:A6";      // Categories

                // Hide rows where the category contains the word "Total"
                for (int row = 1; row <= sheet.Cells.MaxDataRow; row++) // data starts at row index 1 (A2)
                {
                    string category = sheet.Cells[row, 0].StringValue;
                    if (!string.IsNullOrEmpty(category) && category.IndexOf("Total", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        sheet.Cells.Rows[row].IsHidden = true;
                    }
                }

                // Instruct the chart to plot only visible cells
                chart.PlotVisibleCellsOnly = true;

                // Save the workbook
                string outputPath = "ChartFilteredByCategory.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            FilterChartCategories.Run();
        }
    }
}
