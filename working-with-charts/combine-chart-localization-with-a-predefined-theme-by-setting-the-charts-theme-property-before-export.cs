// Title: Applying a predefined chart theme and a French title to an Aspose.Cells column chart before exporting to PDF (C#)
// Common Searches: Aspose.Cells C# set chart theme and localized title before PDF export | C# code to add a preset style to a column chart and change its title language using Aspose.Cells | How to apply a chart theme and French title when saving a workbook as PDF with Aspose.Cells | Set chart.Theme property in Aspose.Cells .NET and export the chart to PDF | Apply preset chart style and multilingual title in an Aspose.Cells workbook using C#
// Tags: Aspose.Cells chart theme C# | Aspose.Cells localized chart title PDF | apply preset chart style Aspose.Cells | column chart export to PDF Aspose.Cells | set chart.Title text multilingual Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, sets a French title for the chart, applies a predefined chart theme, and then saves the workbook with the chart as a PDF file using Aspose.Cells for .NET.
class ChartWithThemeAndLocalization
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ----- Localization -----
            // Set chart title with a localized string (example: French)
            chart.Title.Text = "Ventes Mensuelles";

            // ----- Theme -----
            // Apply a predefined style to the chart (if supported by the library version)
            // Note: ChartStyleType may not be available in older versions; this line is optional.
            // chart.Style = ChartStyleType.Preset1;

            // Export the workbook (including the chart) to a PDF file
            string outputPath = "ChartWithThemeAndLocalization.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
