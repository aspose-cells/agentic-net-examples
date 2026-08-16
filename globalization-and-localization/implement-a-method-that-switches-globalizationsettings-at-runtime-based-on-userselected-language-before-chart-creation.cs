// Title: Runtime Switch of GlobalizationSettings for Localized Charts in Aspose.Cells for .NET
// Description: This example reads a language code from the command line, creates a Workbook, builds a SettableChartGlobalizationSettings object with culture‑specific strings, assigns it to workbook.Settings.GlobalizationSettings, adds sample data, generates a column chart, and saves the file. The chart titles, series names, and legend entries appear in the selected language because the globalization settings are applied before the chart is created.
// Keywords: Aspose.Cells | .NET | chart globalization | runtime localization | SettableChartGlobalizationSettings | Excel chart language | globalization settings | multi‑language Excel | chart title localization | culture code
// Common Searches: Aspose.Cells change chart language at runtime | set chart globalization settings programmatically | localize Excel chart titles with Aspose.Cells | dynamic chart localization .NET | apply SettableGlobalizationSettings before creating chart
// Developer Intent: The developer needs to modify the workbook’s GlobalizationSettings at runtime based on a user‑selected locale so that chart labels are generated in the appropriate language before the chart is added.
// Use Cases: Produce Excel reports with charts that automatically display titles, series names, and legends in the end‑user’s language. | Integrate multi‑language chart generation into a web or desktop application by passing a locale identifier to the chart‑creation routine. | Extend the language‑mapping method to support additional cultures and reuse it for any chart type (column, line, pie, etc.) across the solution.
// AI Prompts: Add Japanese and Chinese entries to GetChartSettingsForLanguage and demonstrate the updated chart localization. | Refactor the switch‑based language mapping to use a dictionary of culture codes and string resources, eliminating the hard‑coded switch. | Show how to retrieve existing GlobalizationSettings from a workbook, modify only the ChartSettings, and keep other localization options unchanged.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example reads a language code from the command line, creates a Workbook, builds a SettableChartGlobalizationSettings object with culture‑specific strings, assigns it to workbook.Settings.GlobalizationSettings, adds sample data, generates a column chart, and saves the file. The chart titles, series names, and legend entries appear in the selected language because the globalization settings are applied before the chart is created.
public class ChartLocalizationHelper
{
    // Entry point required for console application
    public static void Main(string[] args)
    {
        try
        {
            // Determine language code and output path (use defaults if not provided)
            string languageCode = args.Length > 0 ? args[0] : "en";
            string outputPath = args.Length > 1 ? args[1] : "LocalizedChart.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Create the workbook with localized chart settings
            CreateChartWithLocalization(languageCode, outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Creates a workbook, applies language‑specific chart globalization settings,
    // adds a simple column chart and saves the file.
    public static void CreateChartWithLocalization(string languageCode, string outputPath)
    {
        try
        {
            // ---------- Create ----------
            Workbook workbook = new Workbook();                     // create a new workbook
            Worksheet sheet = workbook.Worksheets[0];               // get the first worksheet

            // ---------- Apply Globalization Settings ----------
            // Build chart globalization settings based on the selected language.
            SettableChartGlobalizationSettings chartSettings = GetChartSettingsForLanguage(languageCode);

            // Wrap the chart settings into a SettableGlobalizationSettings instance
            // and assign it to the workbook's globalization settings.
            SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings
            {
                ChartSettings = chartSettings
            };
            workbook.Settings.GlobalizationSettings = globalSettings;

            // ---------- Populate Sample Data ----------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Series");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // ---------- Create Chart ----------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ---------- Save ----------
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Propagate exception to caller after logging
            Console.Error.WriteLine($"Failed to create chart: {ex.Message}");
            throw;
        }
    }

    // Returns a SettableChartGlobalizationSettings instance populated with
    // language‑specific strings. Extend this method with additional languages as needed.
    private static SettableChartGlobalizationSettings GetChartSettingsForLanguage(string languageCode)
    {
        var settings = new SettableChartGlobalizationSettings();

        switch (languageCode.ToLowerInvariant())
        {
            case "en": // English (default)
                settings.SetChartTitleName("Chart Title");
                settings.SetSeriesName("Series");
                settings.SetLegendIncreaseName("Increase");
                settings.SetLegendDecreaseName("Decrease");
                settings.SetOtherName("Other");
                break;

            case "fr": // French
                settings.SetChartTitleName("Titre du graphique");
                settings.SetSeriesName("Série");
                settings.SetLegendIncreaseName("Augmenter");
                settings.SetLegendDecreaseName("Diminuer");
                settings.SetOtherName("Autre");
                break;

            case "de": // German
                settings.SetChartTitleName("Diagrammtitel");
                settings.SetSeriesName("Serie");
                settings.SetLegendIncreaseName("Zunahme");
                settings.SetLegendDecreaseName("Abnahme");
                settings.SetOtherName("Andere");
                break;

            case "es": // Spanish
                settings.SetChartTitleName("Título del gráfico");
                settings.SetSeriesName("Serie");
                settings.SetLegendIncreaseName("Incrementar");
                settings.SetLegendDecreaseName("Disminuir");
                settings.SetOtherName("Otro");
                break;

            default: // Fallback to English
                settings.SetChartTitleName("Chart Title");
                settings.SetSeriesName("Series");
                settings.SetLegendIncreaseName("Increase");
                settings.SetLegendDecreaseName("Decrease");
                settings.SetOtherName("Other");
                break;
        }

        return settings;
    }
}
