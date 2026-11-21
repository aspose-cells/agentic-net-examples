// Title: How to add bold 12‑point titles to the X and Y axes of an Excel chart with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set bold axis title font size 12 for Excel chart | Add X axis and Y axis titles with formatting in Aspose.Cells chart | Aspose.Cells chart axis title customization example in .NET | How to format axis labels in a column chart using Aspose.Cells C# | Set category and value axis titles programmatically with Aspose.Cells
// Tags: format chart axis titles Aspose.Cells C# | bold axis title font size 12 Aspose.Cells | set X and Y axis titles Excel chart .NET | chart axis title customization Aspose.Cells | column chart axis labeling C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, inserts a column chart with sample data, and applies bold 12‑point titles to both the category (X) and value (Y) axes using Aspose.Cells for .NET before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a sample chart to the worksheet (position and size are arbitrary)
            // Add returns the index of the newly created chart; retrieve the chart object using that index
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // ----- Sample data for the chart (optional) -----
            // Populate some data so the chart has something to display
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Set the chart data range
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ----- Add axis titles with formatting -----
            // X (Category) axis title
            chart.CategoryAxis.Title.Text = "X Axis Title";
            chart.CategoryAxis.Title.Font.IsBold = true;   // Bold formatting
            chart.CategoryAxis.Title.Font.Size = 12;       // 12‑point font size

            // Y (Value) axis title
            chart.ValueAxis.Title.Text = "Y Axis Title";
            chart.ValueAxis.Title.Font.IsBold = true;      // Bold formatting
            chart.ValueAxis.Title.Font.Size = 12;          // 12‑point font size

            // Save the workbook (lifecycle rule)
            workbook.Save("ChartWithAxisTitles.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
