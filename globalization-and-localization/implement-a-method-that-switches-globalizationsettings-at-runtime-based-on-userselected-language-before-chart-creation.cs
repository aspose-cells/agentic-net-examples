// Title: How to switch workbook globalization settings at runtime for language‑specific chart formatting using Aspose.Cells for .NET
// Common Searches: set workbook cultureinfo dynamically before creating an Excel chart with Aspose.Cells C# | change Aspose.Cells globalization settings based on user selected language code | apply language specific number and date formats to charts in Aspose.Cells .NET | runtime localization of Excel chart data using Aspose.Cells workbook settings | C# example for switching CultureInfo for chart generation in Aspose.Cells
// Tags: runtime workbook cultureinfo Aspose.Cells | chart localization Aspose.Cells .NET | language specific formatting Excel chart | globalizationsettings per language Aspose.Cells | excel chart cultureinfo switch C#

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The CreateChart method accepts a language code, assigns Workbook.Settings.CultureInfo to a new CultureInfo for that language, populates sample data, adds a column chart linked to the data range, sets a title, and saves the workbook. It also ensures the output directory exists and handles exceptions, enabling language‑aware number and date formatting in generated charts.
    public class ChartGenerator
    {
        // Switches globalization settings based on user language and creates a chart.
        public void CreateChart(string languageCode, string outputPath)
        {
            try
            {
                // Ensure the output directory exists.
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Initialize a new workbook.
                Workbook workbook = new Workbook();

                // Switch globalization settings at runtime.
                // This influences number/date formatting in the chart.
                workbook.Settings.CultureInfo = new CultureInfo(languageCode);

                // Add sample data.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(1200.5);
                sheet.Cells["B3"].PutValue(1500.75);
                sheet.Cells["B4"].PutValue(1100.25);

                // Add a column chart.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart data source.
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set chart title.
                chart.Title.Text = "Sales";

                // Save the workbook.
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating chart: {ex.Message}");
                // Optionally rethrow or handle specific exceptions.
            }
        }
    }

    public class Program
    {
        // Entry point required for the application.
        public static void Main(string[] args)
        {
            try
            {
                string languageCode = "en-US"; // Example language code.
                string outputPath = Path.Combine(Environment.CurrentDirectory, "ChartOutput.xlsx");

                ChartGenerator generator = new ChartGenerator();
                generator.CreateChart(languageCode, outputPath);

                Console.WriteLine($"Chart workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
