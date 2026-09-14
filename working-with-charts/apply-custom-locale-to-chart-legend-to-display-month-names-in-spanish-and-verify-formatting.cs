// Title: Apply a custom Spanish PivotGlobalizationSettings to localize chart legend and category labels in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that derives a class from PivotGlobalizationSettings to return Spanish month abbreviations, fills a worksheet with those months and sample sales data, creates a column chart that uses the months as category labels, positions the legend on the right, and saves the workbook as an XLSX file. | Show how to output the Spanish month names and a confirmation message to the console before saving, to verify that the localization is applied to both the worksheet and the chart.
// Common Searches: asp.net aspocells custom locale for chart legend spanish month names | c# Aspose.Cells set chart category axis labels to Spanish months | how to override PivotGlobalizationSettings for Spanish month abbreviations in Aspose.Cells | save Excel file with localized month names in chart using Aspose.Cells .NET | verify chart legend text after applying custom globalization settings in C#
// Tags: Spanish month globalization settings Aspose.Cells | C# create column chart with localized month labels | right-aligned chart legend Aspose.Cells | populate worksheet with month abbreviations Aspose.Cells | export workbook to XLSX using Aspose.Cells C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Settings;

namespace AsposeCellsChartLocaleDemo
{
    // Custom PivotGlobalizationSettings that returns Spanish month names
    // The example defines a SpanishPivotGlobalizationSettings class that overrides month‑name methods, writes Spanish month abbreviations and sample sales figures to a worksheet, builds a column chart using those months as category labels, aligns the legend to the right, prints verification messages to the console, and saves the result as SpanishMonthChart.xlsx.
    public class SpanishPivotGlobalizationSettings : PivotGlobalizationSettings
    {
        // Override to provide short month names in Spanish
        public override string[] GetShortTextOf12Months()
        {
            return new string[]
            {
                "Ene", "Feb", "Mar", "Abr", "May", "Jun",
                "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"
            };
        }

        // Optionally override full month names as well
        public override string GetTextOfMonths()
        {
            return "Meses";
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Instantiate the custom Spanish month settings
                SpanishPivotGlobalizationSettings spanishSettings = new SpanishPivotGlobalizationSettings();

                // Retrieve Spanish short month names
                string[] spanishMonths = spanishSettings.GetShortTextOf12Months();

                // Populate worksheet with month names and sample data
                sheet.Cells["A1"].PutValue("Mes");          // Header for month column
                sheet.Cells["B1"].PutValue("Ventas");       // Header for data column

                for (int i = 0; i < spanishMonths.Length; i++)
                {
                    // Column A: month name
                    sheet.Cells[i + 1, 0].PutValue(spanishMonths[i]);
                    // Column B: sample sales value
                    sheet.Cells[i + 1, 1].PutValue(100 + i * 20);
                }

                // Create a column chart that uses the month names as categories
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series (values) and categories (months)
                chart.NSeries.Add("B2:B13", true);
                chart.NSeries.CategoryData = "A2:A13";

                // Set chart title
                chart.Title.Text = "Ventas Mensuales";

                // Configure legend position (visibility is true by default)
                chart.Legend.Position = LegendPositionType.Right;

                // Output verification info to console
                Console.WriteLine("Spanish month names used in the worksheet:");
                foreach (string month in spanishMonths)
                {
                    Console.WriteLine(month);
                }

                Console.WriteLine("\nChart created with category axis showing Spanish month names.");
                Console.WriteLine("Legend displays the series name derived from the data range.");

                // Save the workbook
                string outputPath = "SpanishMonthChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
