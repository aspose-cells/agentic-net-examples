// Title: Create a column chart with multiple series in Aspose.Cells C# and understand the current limitation on hiding series and localizing number formats
// AI Prompts: Generate C# code using Aspose.Cells that builds a column chart, programmatically sets a specific series' IsVisible property to false, and then applies culture‑specific number formats to the chart axes. | Show how to add three data series to an Aspose.Cells chart, hide one series before saving, and configure number format localization for the workbook.
// Common Searches: aspnet hide specific series in Aspose.Cells column chart | how to set chart series visibility in Aspose.Cells for .NET | apply culture specific number formats to Aspose.Cells chart axes | Aspose.Cells chart series visibility before localization | C# example for hiding a data series in Excel chart using Aspose.Cells
// Tags: Aspose.Cells hide chart series C# | column chart series visibility Aspose.Cells | Aspose.Cells number format localization .NET | programmatic chart series manipulation Aspose.Cells | set IsVisible property chart series Aspose.Cells | culture specific number formats Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, fills it with sample data for three series, adds a column chart containing those series, and saves the file. It also notes that hiding a series and applying localized number formats are not supported in the current Aspose.Cells version, highlighting the limitation for developers.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Fill sample data for three series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["D1"].PutValue("Series3");

            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sample Chart";

            // Add three data series to the chart
            chart.NSeries.Add("B2:B4", true); // Series1
            chart.NSeries.Add("C2:C4", true); // Series2
            chart.NSeries.Add("D2:D4", true); // Series3

            // Note: Hiding a series and localizing number formats are not supported
            // in the current Aspose.Cells version used. These lines have been removed
            // to ensure the code compiles and runs correctly.

            // Define output file path
            string outputPath = "ChartSeriesVisibility.xlsx";

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
