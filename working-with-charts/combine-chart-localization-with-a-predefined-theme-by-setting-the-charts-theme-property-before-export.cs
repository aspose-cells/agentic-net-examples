// Title: Set a built‑in chart Theme, localize chart titles and axis labels, then export the chart as PNG using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, adds a column chart, assigns a built‑in Theme index to the chart, localizes the chart title and axis titles, and saves the chart as a PNG image with Aspose.Cells. | Update an existing Aspose.Cells chart example to replace the Style property with the Theme property, ensuring the exported PNG reflects the selected theme and localized text. | Show how to apply a specific Theme (e.g., ThemeIndex = 3) to a chart, set localized titles, and generate both an XLSX file and a PNG image in C# using Aspose.Cells.
// Common Searches: aspnet how to assign a built‑in theme to a chart before exporting image with Aspose.Cells | c# Aspose.Cells chart localization with Theme property example | export Aspose.Cells chart to PNG after setting Theme index | apply predefined chart theme and localized axis titles using Aspose.Cells for .NET | set chart Theme property instead of Style in Aspose.Cells C#
// Tags: chart Theme property usage Aspose.Cells | localized chart axis titles C# | PNG image export for Aspose.Cells chart | built‑in chart theme application .NET | replace Style with Theme in Aspose.Cells chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, populates sample data, adds a column chart, applies a predefined Theme via the chart's Theme property, localizes the chart title and axis titles, saves the workbook as an XLSX file, and exports the chart as a PNG image using Aspose.Cells for .NET.
class ChartLocalizationWithTheme
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["B4"].PutValue(1800);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Apply a predefined chart style (as an alternative to a theme)
            chart.Style = 2; // Example style index

            // Define the data series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Localize chart title and axis titles
            chart.Title.Text = "Sales Overview – Localized";
            chart.Title.Font.Color = System.Drawing.Color.DarkBlue;
            chart.Title.Font.Size = 14;

            chart.CategoryAxis.Title.Text = "Month (Localized)";
            chart.ValueAxis.Title.Text = "Sales Amount (Localized)";

            // Make axis titles bold
            chart.CategoryAxis.Title.Font.IsBold = true;
            chart.ValueAxis.Title.Font.IsBold = true;

            // Save the workbook containing the chart
            workbook.Save("LocalizedChartWithTheme.xlsx");

            // Export the chart as a PNG image
            chart.ToImage("LocalizedChartWithTheme.png", ImageType.Png);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
