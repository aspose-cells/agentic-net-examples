// Title: Apply a custom RGB theme palette to each series of a multi‑series column chart with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates an Excel workbook, adds a clustered column chart, and assigns a unique RGB foreground color to each series using Aspose.Cells. | Show how to use the Aspose.Cells chart API to set custom colors on NSeries objects and save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# set individual series colors in a column chart | How to change the foreground color of each series in an Excel chart using Aspose.Cells .NET | Apply custom RGB palette to a multi‑series chart with Aspose.Cells | C# Aspose.Cells example for coloring chart series separately | Assign custom colors to NSeries in Aspose.Cells chart
// Tags: set series foreground color Aspose.Cells | custom RGB palette chart Aspose.Cells .NET | apply theme colors to Excel chart series | multi-series column chart color customization Aspose.Cells | Aspose.Cells NSeries color API

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills it with sample data, adds a clustered column chart, and applies distinct RGB colors (Tomato, MediumSeaGreen, DodgerBlue) to each of the three series before saving the file as CustomPaletteChart.xlsx.
class ApplyCustomThemePalette
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a multi‑series chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["D1"].PutValue("Series 3");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(45);

            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);
            sheet.Cells["D5"].PutValue(42);

            // Add a clustered column chart (ChartType.Column is the clustered column)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (including categories and series)
            chart.NSeries.Add("B2:D5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Apply custom colors to each series
            chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(255, 99, 71);   // Tomato
            chart.NSeries[1].Area.ForegroundColor = Color.FromArgb(60, 179, 113); // MediumSeaGreen
            chart.NSeries[2].Area.ForegroundColor = Color.FromArgb(30, 144, 255); // DodgerBlue

            // Save the workbook to a file
            string outputPath = "CustomPaletteChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
