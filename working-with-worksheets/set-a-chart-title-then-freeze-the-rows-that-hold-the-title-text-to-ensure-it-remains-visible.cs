// Title: How to add a column chart with a formatted title and freeze the title rows in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart title font color and freeze top rows | Freeze panes above chart title in Excel using Aspose.Cells .NET | Create column chart with title and keep title visible while scrolling Aspose.Cells | C# Aspose.Cells example for adding chart title and freezing rows | How to use FreezePanes method to lock chart title rows in Aspose.Cells
// Tags: add column chart with data series Aspose.Cells | format chart title font Aspose.Cells | apply freeze panes to top rows Aspose.Cells | freeze rows using cell reference A2 Aspose.Cells | export workbook to xlsx format Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTitleFreeze
{
    // The C# sample creates a new workbook, fills it with sample data, adds a column chart, formats a visible centered title with Calibri 14 dark‑blue font, freezes the rows above cell A2 so the title stays in view while scrolling, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Populate sample data for the chart
                // -------------------------------------------------
                Cells cells = sheet.Cells;
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("Alpha");
                cells["B2"].PutValue(120);
                cells["A3"].PutValue("Beta");
                cells["B3"].PutValue(150);
                cells["A4"].PutValue("Gamma");
                cells["B4"].PutValue(180);

                // -------------------------------------------------
                // 2. Add a chart to the worksheet
                // -------------------------------------------------
                // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // -------------------------------------------------
                // 3. Set the chart title
                // -------------------------------------------------
                chart.Title.Text = "Quarterly Sales Overview";
                chart.Title.IsVisible = true;
                chart.Title.Font.Name = "Calibri";
                chart.Title.Font.Size = 14;
                chart.Title.Font.Color = Color.DarkBlue;
                chart.Title.TextHorizontalAlignment = TextAlignmentType.Center;

                // -------------------------------------------------
                // 4. Freeze the rows that contain the title text
                // -------------------------------------------------
                // Freeze the top row (row index 0) so the title stays visible while scrolling.
                // Using the overload that accepts a cell name: freeze panes above cell A2.
                sheet.FreezePanes("A2", 1, 0);

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                string outputPath = "ChartTitleFreezeDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Output any errors that occur during processing
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
