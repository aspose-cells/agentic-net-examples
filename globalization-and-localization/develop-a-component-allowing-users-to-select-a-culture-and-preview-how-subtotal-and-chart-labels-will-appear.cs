// Title: Aspose.Cells .NET – Culture‑aware workbook with custom subtotal and chart label globalization
// Description: A C# console sample that prompts for a culture code (e.g., en‑US, fr‑FR), applies the CultureInfo to a Workbook, defines custom subtotal text and chart legends via SettableGlobalizationSettings and SettableChartGlobalizationSettings, builds sample data, inserts a column chart, displays the localized labels in the console, and saves the file as CulturePreview.xlsx.
// Keywords: Aspose.Cells | .NET | C# | CultureInfo | globalization | localization | subtotal label customization | chart legend localization | SettableGlobalizationSettings | SettableChartGlobalizationSettings | Excel workbook culture | subtotal function | chart title localization | console example
// Common Searches: Aspose.Cells change subtotal label language | set workbook culture programmatically Aspose.Cells .NET | customize chart legend text based on culture Aspose.Cells | preview Excel chart labels for different locales | globalization settings example Aspose.Cells C#
// Developer Intent: Provide an interactive tool that lets users pick a locale and instantly see how subtotal and chart labels are rendered with that culture in an Excel file.
// Use Cases: Create localized financial statements where subtotal rows show culture‑specific terminology | Demonstrate how chart titles, series names, and legend entries adapt to different locales for training or documentation | Integrate a quick‑check console utility into CI pipelines to verify globalization settings before release
// AI Prompts: Generate a reusable method that parses a culture string, falls back to invariant culture on error, and assigns it to Workbook.Settings.CultureInfo. | Show code to read the custom total name for the Sum function after configuring SettableGlobalizationSettings and output it to the console. | Explain how to extend the sample to define custom labels for Average, Count, and Max functions in the subtotal globalization configuration. | Provide a step‑by‑step guide to export the workbook as PDF while preserving the localized chart labels.

using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

namespace AsposeCellsCulturePreview
{
    // A C# console sample that prompts for a culture code (e.g., en‑US, fr‑FR), applies the CultureInfo to a Workbook, defines custom subtotal text and chart legends via SettableGlobalizationSettings and SettableChartGlobalizationSettings, builds sample data, inserts a column chart, displays the localized labels in the console, and saves the file as CulturePreview.xlsx.
    class Program
    {
        static void Main()
        {
            // Prompt user for a culture name (e.g., "en-US", "fr-FR")
            Console.Write("Enter culture name (e.g., en-US, fr-FR): ");
            string cultureName = Console.ReadLine();

            // Validate and create CultureInfo
            CultureInfo culture;
            try
            {
                culture = new CultureInfo(cultureName);
            }
            catch (CultureNotFoundException)
            {
                Console.WriteLine("Invalid culture. Using invariant culture.");
                culture = CultureInfo.InvariantCulture;
            }

            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Apply the selected culture to the workbook
            workbook.Settings.CultureInfo = culture;

            // -------------------- Globalization Settings --------------------
            // Create chart globalization settings and customize labels
            SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
            chartSettings.SetSeriesName("Custom Series");
            chartSettings.SetChartTitleName("Custom Chart Title");
            chartSettings.SetLegendIncreaseName("Increase");
            chartSettings.SetLegendDecreaseName("Decrease");
            chartSettings.SetLegendTotalName("Total");

            // Create settable globalization settings for subtotal and other texts
            SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings();

            // Customize subtotal label for Sum function
            globalSettings.SetTotalName(ConsolidationFunction.Sum, "Custom Sum Total");

            // Assign the chart settings to the globalization settings
            globalSettings.ChartSettings = chartSettings;

            // Apply globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalSettings;

            // -------------------- Sample Data --------------------
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");

            // Data rows
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("C");
            cells["B4"].PutValue(30);

            // -------------------- Subtotal --------------------
            // Apply subtotal on the data range; the subtotal label will use the customized total name
            CellArea area = CellArea.CreateCellArea(0, 0, 4, 1); // A1:B5
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // -------------------- Chart --------------------
            // Create a simple column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add series data
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set chart title (will be displayed using the customized title name)
            chart.Title.Text = "Demo Chart";

            // -------------------- Preview Output --------------------
            Console.WriteLine();
            Console.WriteLine("=== Preview ===");
            Console.WriteLine($"Culture: {culture.Name}");
            Console.WriteLine($"Subtotal label for Sum: {globalSettings.GetTotalName(ConsolidationFunction.Sum)}");
            Console.WriteLine($"Chart series default name: {chartSettings.GetSeriesName()}");
            Console.WriteLine($"Chart title default name: {chartSettings.GetChartTitleName()}");
            Console.WriteLine($"Chart legend increase name: {chartSettings.GetLegendIncreaseName()}");
            Console.WriteLine($"Chart legend decrease name: {chartSettings.GetLegendDecreaseName()}");
            Console.WriteLine();

            // Save the workbook (lifecycle rule: save)
            string outputPath = "CulturePreview.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
