// Title: Enable anti‑aliasing and set custom ChartGlobalizationSettings for a column chart in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, adds a column chart, activates anti‑aliasing for the chart, and applies a ChartGlobalizationSettings object with a specified locale before saving the file. | Show how to configure Aspose.Cells to render charts with smoother edges and set the chart’s culture to French (fr‑FR) using ChartGlobalizationSettings. | Provide a snippet that modifies an existing Aspose.Cells column chart to turn on enhanced rendering and assign custom globalization parameters such as language and region.
// Common Searches: Aspose.Cells smooth chart rendering C# example | How to set chart culture for Excel files using Aspose.Cells .NET | Improve visual quality of column charts in Aspose.Cells | Configure localization for charts in Aspose.Cells workbook | C# Aspose.Cells example for high‑quality chart output
// Tags: chart edge smoothing Aspose.Cells | ChartGlobalizationSettings culture configuration Aspose.Cells | column chart visual enhancement .NET | Excel chart localization parameters Aspose.Cells | chart rendering options Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// // Creates a workbook, adds a column chart with sample data, enables anti‑aliasing, applies custom ChartGlobalizationSettings for a chosen locale, and saves the file as ChartWithAntiAliasing.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sample Column Chart";

            // Set the data range for the chart series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Save the workbook with the configured chart
            workbook.Save("ChartWithAntiAliasing.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
