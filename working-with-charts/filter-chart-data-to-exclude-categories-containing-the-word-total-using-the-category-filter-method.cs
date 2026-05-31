using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (Category – Column A, Value – Column B)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            sheet.Cells["A2"].PutValue("North");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("South");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("East Total");   // should be excluded
            sheet.Cells["B4"].PutValue(200);
            sheet.Cells["A5"].PutValue("West");
            sheet.Cells["B5"].PutValue(130);
            sheet.Cells["A6"].PutValue("Central Total"); // should be excluded
            sheet.Cells["B6"].PutValue(170);

            // Add a column chart that uses the whole data range
            int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B6", true);          // Values
            chart.NSeries.CategoryData = "A2:A6";      // Categories

            // Hide rows whose category contains the word "Total"
            for (int row = 1; row <= sheet.Cells.MaxDataRow; row++) // start from row 1 (A2)
            {
                string category = sheet.Cells[row, 0].StringValue; // column A
                if (!string.IsNullOrEmpty(category) &&
                    category.IndexOf("Total", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Use Cells.Rows collection to hide the row
                    sheet.Cells.Rows[row].IsHidden = true;
                }
            }

            // Instruct the chart to plot only visible cells
            chart.PlotVisibleCellsOnly = true;

            // Save the workbook (ensure the directory exists)
            string outputPath = "ChartFilteredByCategory.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}