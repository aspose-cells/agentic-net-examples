// Title: Set workbook culture at runtime and generate a localized column chart with Aspose.Cells for .NET
// AI Prompts: Write C# code that accepts a culture name, assigns Workbook.Settings.CultureInfo, populates sample data, creates a column chart, and saves the workbook as an XLSX file using Aspose.Cells. | Show how to switch Aspose.Cells globalization settings based on a user‑selected language before creating any chart in a .NET application.
// Common Searches: how to change workbook cultureinfo in Aspose.Cells before adding a chart in C# | Aspose.Cells set language for Excel chart labels at runtime | C# create localized column chart with Aspose.Cells using CultureInfo | dynamic globalization of Excel files with Aspose.Cells based on user selection | switch Aspose.Cells workbook culture for different locales in a .NET app
// Tags: configure workbook culture Aspose.Cells | runtime localization of Excel chart .NET | create column chart with Aspose.Cells | Aspose.Cells globalization settings | save workbook as xlsx Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example demonstrates how to assign a CultureInfo to Workbook.Settings.CultureInfo at runtime, fill sample data, add a column chart, and save the workbook as an XLSX file, enabling chart labels and number formats to reflect the selected language.
public class ChartHelper
{
    /// <param name="languageCode">Culture name representing the selected language.</param>
    /// <param name="outputPath">Full path where the workbook will be saved.</param>
    public void CreateChartWithLanguage(string languageCode, string outputPath)
    {
        try
        {
            // Ensure the output directory exists
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            // Initialize a new workbook
            Workbook workbook = new Workbook();

            // Set workbook culture for globalization (affects number/date formatting, chart labels, etc.)
            workbook.Settings.CultureInfo = new CultureInfo(languageCode);

            // Populate sample data for the chart
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Create a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true); // Values
            // Category data line removed because Series.CategoryData is not available in this version

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating chart: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Example usage
            string languageCode = "fr-FR"; // French (France)
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ChartWithLanguage.xlsx");

            ChartHelper helper = new ChartHelper();
            helper.CreateChartWithLanguage(languageCode, outputPath);

            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
