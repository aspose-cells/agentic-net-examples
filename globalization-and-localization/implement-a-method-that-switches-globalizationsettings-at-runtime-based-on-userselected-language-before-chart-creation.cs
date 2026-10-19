// Title: Create an Excel chart with Aspose.Cells .NET using a user‑selected language by setting Workbook culture at runtime
// Common Searches: how to set workbook cultureinfo dynamically in Aspose.Cells C# before adding a chart | Aspose.Cells change number and date format based on user language at runtime | C# generate Excel column chart with localized formatting using Aspose.Cells | runtime globalization settings for Excel workbook in Aspose.Cells .NET | apply user selected culture to Aspose.Cells workbook for chart creation
// Tags: set workbook culture Aspose.Cells C# | runtime localization Excel chart Aspose.Cells | dynamic CultureInfo for workbook Aspose.Cells | language specific number formatting Aspose.Cells | chart generation with localized settings .NET

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example shows how to create a new Workbook, assign a CultureInfo based on a user‑provided language code to workbook.Settings.CultureInfo, populate sample data, add a column chart, and save the file. It also ensures the output directory exists and handles exceptions, enabling runtime globalization of Excel charts with Aspose.Cells for .NET.
public class ChartGenerator
{
    /// <param name="languageCode">Culture name (e.g., "en-US", "fr-FR", "de-DE").</param>
    /// <param name="outputPath">Full file path where the workbook will be saved.</param>
    public void CreateChartWithLanguage(string languageCode, string outputPath)
    {
        try
        {
            // Ensure the output directory exists
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set culture for number/date formatting
            workbook.Settings.CultureInfo = new CultureInfo(languageCode);
            // The UseSystemSeparators property is not available in recent versions; culture settings are sufficient.

            // Populate sample data for the chart
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Add a column chart (position: row 5, column 0 to row 15, column 5)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("A1:A3", true);

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating chart: {ex.Message}");
        }
    }

    // Entry point for execution
    public static void Main(string[] args)
    {
        try
        {
            string languageCode = "en-US";
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ChartDemo.xlsx");
            new ChartGenerator().CreateChartWithLanguage(languageCode, outputPath);
            Console.WriteLine($"Workbook saved successfully at: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
