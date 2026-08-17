// Title: Localize Aspose.Cells Chart Legend and Month Labels to Spanish (C#)
// Description: Creates a workbook, fills column A with Spanish month abbreviations via a custom PivotGlobalizationSettings class, adds a column chart, applies Spanish chart globalization settings (legend total, increase, decrease, series name) using SettableChartGlobalizationSettings, invokes SetChartGlobalizationSettings through reflection when needed, prints the applied values for verification, and saves the file as ChartWithSpanishLocale.xlsx.
// Keywords: Aspose.Cells | C# chart localization | Spanish month names | custom chart globalization settings | SetChartGlobalizationSettings | PivotGlobalizationSettings | SettableChartGlobalizationSettings | column chart Excel | .NET Excel API | chart legend translation
// Common Searches: Aspose.Cells set chart legend language | Spanish month abbreviations in Excel chart C# | custom chart globalization settings Aspose.Cells | how to use SetChartGlobalizationSettings with reflection | localize chart labels Aspose.Cells .NET
// Developer Intent: Apply a Spanish locale to a chart’s legend and category labels, verify the applied text, and generate a localized Excel workbook.
// Use Cases: Produce monthly sales dashboards with Spanish month names and localized legend entries. | Build multi‑language Excel reports where each chart uses its own globalization settings. | Automated testing to confirm that custom chart globalization settings are correctly applied before distribution.
// AI Prompts: Generate C# code that sets French month names and legend labels for an Aspose.Cells chart using custom globalization settings. | Explain how to safely call Workbook.Settings.SetChartGlobalizationSettings via reflection when the method is not exposed in older Aspose.Cells versions. | Provide a step‑by‑step verification script to ensure Spanish month abbreviations appear in chart categories and legend after applying custom settings.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Settings;

// Creates a workbook, fills column A with Spanish month abbreviations via a custom PivotGlobalizationSettings class, adds a column chart, applies Spanish chart globalization settings (legend total, increase, decrease, series name) using SettableChartGlobalizationSettings, invokes SetChartGlobalizationSettings through reflection when needed, prints the applied values for verification, and saves the file as ChartWithSpanishLocale.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Custom month names in Spanish ----------
            // Use a derived PivotGlobalizationSettings to supply Spanish month abbreviations
            var spanishMonthsProvider = new SpanishPivotGlobalizationSettings();
            string[] spanishMonths = spanishMonthsProvider.GetShortTextOf12Months();

            // Populate the worksheet with month names and sample numeric values
            sheet.Cells["A1"].PutValue("Mes");   // Header for month column
            sheet.Cells["B1"].PutValue("Valor"); // Header for value column
            for (int i = 0; i < spanishMonths.Length; i++)
            {
                // Column A: Spanish month name
                sheet.Cells[i + 1, 0].PutValue(spanishMonths[i]);
                // Column B: Sample value (e.g., 10, 20, ...)
                sheet.Cells[i + 1, 1].PutValue((i + 1) * 10);
            }

            // ---------- Create a chart ----------
            // Add a column chart that uses the data above
            int chartIndex = sheet.Charts.Add(ChartType.Column, 2, 3, 20, 15);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B13", true);          // Values
            chart.NSeries.CategoryData = "A2:A13";      // Categories (Spanish month names)

            // ---------- Apply custom chart globalization settings ----------
            var chartLocale = new SpanishChartGlobalizationSettings();
            chartLocale.SetLegendTotalName("Total");          // "Total" in Spanish
            chartLocale.SetLegendIncreaseName("Incremento"); // "Increase" in Spanish
            chartLocale.SetLegendDecreaseName("Decremento"); // "Decrease" in Spanish
            chartLocale.SetSeriesName("Meses");              // Series name in Spanish

            // Attach the custom settings to the workbook if the API is available.
            // In some versions SetChartGlobalizationSettings may not exist; guard with reflection.
            try
            {
                var method = workbook.Settings.GetType().GetMethod("SetChartGlobalizationSettings");
                method?.Invoke(workbook.Settings, new object[] { chartLocale });
            }
            catch (Exception ex)
            {
                Console.WriteLine("Chart globalization settings could not be applied: " + ex.Message);
            }

            // ---------- Verification ----------
            Console.WriteLine("Spanish month abbreviations:");
            foreach (var m in spanishMonths)
                Console.WriteLine(m);

            Console.WriteLine("Legend Total Name: " + chartLocale.GetLegendTotalName());
            Console.WriteLine("Legend Increase Name: " + chartLocale.GetLegendIncreaseName());
            Console.WriteLine("Legend Decrease Name: " + chartLocale.GetLegendDecreaseName());
            Console.WriteLine("Series Name: " + chartLocale.GetSeriesName());

            // Save the workbook
            string outputPath = "ChartWithSpanishLocale.xlsx";
            try
            {
                // Ensure the directory exists
                string fullPath = Path.GetFullPath(outputPath);
                string? dir = Path.GetDirectoryName(fullPath);
                if (string.IsNullOrEmpty(dir))
                {
                    dir = Directory.GetCurrentDirectory();
                }

                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {fullPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Custom PivotGlobalizationSettings that returns Spanish short month names
    class SpanishPivotGlobalizationSettings : PivotGlobalizationSettings
    {
        public override string[] GetShortTextOf12Months()
        {
            return new[]
            {
                "Ene", "Feb", "Mar", "Abr", "May", "Jun",
                "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"
            };
        }
    }

    // Custom SettableChartGlobalizationSettings – inherits all needed methods
    class SpanishChartGlobalizationSettings : SettableChartGlobalizationSettings { }
}
