// Title: Localize a chart legend to French using Aspose.Cells in C#
// AI Prompts: Generate C# code that sets the workbook CultureInfo to fr-FR and updates the chart legend text to French with Aspose.Cells. | Provide a step‑by‑step example that creates a column chart, assigns French month labels, customizes the legend title in French, and renders the chart to a PNG image. | Show how to save the workbook containing the French‑localized chart as an XLSX file using Aspose.Cells. | Explain how to verify the legend text after localization by reading the Legend.Text property.
// Common Searches: aspocells set workbook cultureinfo to fr-fr for chart legend | c# change chart legend text to French with Aspose.Cells | render Aspose.Cells chart to PNG after localizing legend | save workbook with French chart labels using Aspose.Cells | verify chart legend language in Aspose.Cells C# example
// Tags: set workbook cultureinfo for chart localization Aspose.Cells | chart legend text localization Aspose.Cells | render chart to png Aspose.Cells | save workbook as xlsx Aspose.Cells | column chart with French month labels Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a workbook, sets its CultureInfo to fr-FR, adds a column chart with French month names, customizes the legend title to French, renders the chart to a PNG image for visual verification, and saves the workbook as an XLSX file.
class ChartLocalizationExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set workbook culture to French (France) for globalization
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Janvier");   // French month names
            sheet.Cells["A3"].PutValue("Février");
            sheet.Cells["A4"].PutValue("Mars");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["B4"].PutValue(1800);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set a custom legend text (localized)
            chart.Legend.Text = "Légende"; // French for "Legend"

            // Verify that the legend text reflects the French text
            string legendText = chart.Legend.Text;
            Console.WriteLine("Legend Text after localization: " + legendText);

            // Render the chart to an image to visually verify language rendering
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions(); // Default format is PNG
            chart.ToImage("LocalizedChart.png", imgOptions);

            // Save the workbook
            workbook.Save("LocalizedChartWorkbook.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
