// Title: C# – Hide “Total” rows to filter chart categories with Aspose.Cells
// Description: Creates a workbook, populates category and value columns, adds a column chart, enables PlotVisibleCellsOnly, hides rows whose category text contains "Total", and saves the file so the chart displays only the visible categories.
// Keywords: Aspose.Cells C# | hide rows chart filter | PlotVisibleCellsOnly | exclude Total categories | Excel chart category filter | Aspose.Cells .NET chart | column chart data filtering | category keyword filter
// Common Searches: Aspose.Cells hide rows containing specific text | filter chart categories by keyword Aspose.Cells | PlotVisibleCellsOnly example C# | exclude total rows from Excel chart using Aspose.Cells | how to hide rows so chart ignores them .NET
// Developer Intent: Remove rows whose category includes the word "Total" so the chart automatically omits those categories while keeping the data in the worksheet.
// Use Cases: Generate a sales report where summary rows (e.g., "East Total") stay in the sheet but are not plotted in the chart. | Create dynamic dashboards that hide specific categories without altering the source data. | Build reusable Excel templates that automatically filter out total rows from visualizations.
// AI Prompts: Write C# code with Aspose.Cells to hide rows where the category column contains "Total" and update a column chart to show only visible categories. | Explain how the PlotVisibleCellsOnly property works and how it can be combined with row hiding to filter chart data in Aspose.Cells. | Provide an Aspose.Cells .NET example that filters chart categories by a keyword without deleting the underlying rows.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCategoryFilter
{
    // Creates a workbook, populates category and value columns, adds a column chart, enables PlotVisibleCellsOnly, hides rows whose category text contains "Total", and saves the file so the chart displays only the visible categories.
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
                // Column A – Category, Column B – Value
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

                // Add a column chart that uses the above data
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B6", true);          // Values
                chart.NSeries.CategoryData = "A2:A6";      // Categories

                // Instruct the chart to plot only visible cells.
                // Rows that are hidden will be ignored by the chart.
                chart.PlotVisibleCellsOnly = true;

                // Hide rows where the category contains the word "Total"
                // This effectively excludes those categories from the chart.
                for (int row = 1; row <= sheet.Cells.MaxDataRow; row++) // start from row 2 (index 1)
                {
                    string category = sheet.Cells[row, 0].StringValue;
                    if (!string.IsNullOrEmpty(category) &&
                        category.IndexOf("Total", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Use Cells.Rows collection to hide the row
                        sheet.Cells.Rows[row].IsHidden = true;
                    }
                }

                // Save the workbook
                workbook.Save("ChartFilteredByCategory.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main()
        {
            Run();
        }
    }
}
