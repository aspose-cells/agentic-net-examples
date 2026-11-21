// Title: Apply a Chinese chart title and a predefined chart style with Aspose.Cells for .NET before exporting to XLSX
// Common Searches: C# Aspose.Cells set chart title in Chinese | apply predefined chart style index using Aspose.Cells before saving workbook | how to combine chart localization and theme in Aspose.Cells .NET | example of setting chart title localization and style with Aspose.Cells | Aspose.Cells export chart with custom style and localized title to XLSX
// Tags: chart title localization Aspose.Cells C# | predefined chart style index Aspose.Cells | export chart with theme Aspose.Cells .NET | Excel chart styling using Aspose.Cells | localized Excel chart example C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills month and sales data, adds a column chart, sets the chart title to a Chinese string, applies a preset style via the Style property, and saves the file as LocalizedChartWithTheme.xlsx using Aspose.Cells for .NET.
class Program
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
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["B4"].PutValue(1800);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Localize chart title (e.g., Chinese)
            chart.Title.Text = "销售额";

            // Apply a predefined chart style (using integer index as ChartStyleType is unavailable)
            chart.Style = 1; // Preset style 1

            // Save the workbook to an Excel file
            string outputPath = "LocalizedChartWithTheme.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
