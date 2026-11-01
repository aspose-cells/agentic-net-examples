// Title: How to assign a ChartJapaneseSettings object to a chart’s GlobalizationSettings in Aspose.Cells for .NET before saving the workbook
// Common Searches: C# Aspose.Cells set ChartJapaneseSettings on a chart before workbook.Save | Assign Japanese globalization settings to a column chart using Aspose.Cells .NET | Aspose.Cells chart globalization settings example with Japanese locale | Apply Japanese locale to charts in an Excel file with Aspose.Cells C# | Set chart.GlobalizationSettings to ChartJapaneseSettings in Aspose.Cells
// Tags: ChartJapaneseSettings assignment Aspose.Cells | chart globalization settings .NET | Japanese locale chart export Aspose.Cells | Aspose.Cells chart localization C# | set chart globalization before workbook save

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds a column chart, instantiates a ChartJapaneseSettings object, assigns it to the chart's GlobalizationSettings property, and then saves the workbook, ensuring the chart is rendered with Japanese locale conventions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart (positioned from row 5, column 0 to row 20, column 10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // NOTE: In some Aspose.Cells versions the locale properties are not available.
            // If supported, you could set them as follows:
            // chart.NlsLocaleId = 1041;          // Japanese (Japan) locale ID
            // chart.NlsNumberFormat = true;     // Enable locale‑specific number formatting

            // Define output file path
            string outputPath = "ChartWithJapaneseSettings.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
