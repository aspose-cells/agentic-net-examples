// Title: Add a chart title with custom font size and color using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart title font size to 14 points | How to change chart title color to dark blue in Aspose.Cells .NET | C# example for adding a visible title to a column chart with Aspose.Cells | Customize chart title alignment and font in Aspose.Cells workbook | Aspose.Cells chart title formatting tutorial for .xlsx files | Set chart title text and style programmatically with Aspose.Cells C#
// Tags: chart title font customization Aspose.Cells | set chart title color .NET | column chart title styling Aspose.Cells | Aspose.Cells chart title alignment | custom chart title appearance C# | Aspose.Cells workbook chart formatting

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills month and sales data, adds a column chart, makes the chart title visible, sets its text, applies a Calibri font of size 14 with dark blue color, centers the title horizontally, and saves the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the chart
            cells["A1"].PutValue("Month");
            cells["B1"].PutValue("Sales");

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] sales = { 12000, 15000, 13000, 17000, 16000 };

            for (int i = 0; i < months.Length; i++)
            {
                cells[i + 1, 0].PutValue(months[i]);   // Column A
                cells[i + 1, 1].PutValue(sales[i]);   // Column B
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B6", true);          // Values
            chart.NSeries.CategoryData = "A2:A6";      // Categories

            // Insert a chart title with custom font size and color
            chart.Title.IsVisible = true;
            chart.Title.Text = "Monthly Sales Report";

            // Customize the title's font
            chart.Title.Font.Name = "Calibri";
            chart.Title.Font.Size = 14;                // Font size in points
            chart.Title.Font.Color = Color.DarkBlue;   // Font color

            // Optional: center the title horizontally
            chart.Title.TextHorizontalAlignment = TextAlignmentType.Center;

            // Save the workbook
            string outputFile = "ChartWithCustomTitle.xlsx";
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
