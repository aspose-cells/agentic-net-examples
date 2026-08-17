// Title: C# Aspose.Cells example: Select a culture to preview pivot subtotal and chart labels
// Description: A console‑based C# sample that asks for a locale (e.g., en‑US, fr‑FR), applies the corresponding CultureInfo to a Workbook, configures custom subtotal texts for a pivot table and custom chart legends via SettablePivotGlobalizationSettings and SettableChartGlobalizationSettings, displays the applied texts and number‑group separator, and saves the result as CulturePreview.xlsx.
// Keywords: Aspose.Cells C# example | culture selection Excel | Workbook.Settings.CultureInfo | SettablePivotGlobalizationSettings | SettableChartGlobalizationSettings | pivot subtotal label localization | chart legend globalization | Excel workbook localization | console utility Aspose.Cells | GitHub Aspose.Cells sample
// Common Searches: how to change pivot subtotal text Aspose.Cells | custom chart labels based on culture Aspose.Cells .NET | apply CultureInfo to Aspose.Cells workbook | preview number formatting in Aspose.Cells console app | Aspose.Cells globalization settings example
// Developer Intent: Provide a ready‑to‑run C# snippet that lets developers choose a locale and instantly see localized pivot subtotal and chart labels before generating the final Excel file.
// Use Cases: Create a locale‑aware reporting tool that previews how pivot subtotals and chart legends appear for different regions. | Generate region‑specific Excel reports with automatically translated chart titles, series names, and legend entries. | Validate number‑format symbols (group separator, decimal separator) for multiple cultures during development.
// AI Prompts: Generate C# code that prompts for a culture string, sets Workbook.Settings.CultureInfo, customizes pivot subtotal and chart globalization texts, prints a preview, and saves the workbook using Aspose.Cells. | Explain the relationship between SettablePivotGlobalizationSettings, SettableChartGlobalizationSettings, and Workbook.Settings.GlobalizationSettings for label localization. | Show how to retrieve and display the current number group separator from Workbook.Settings.CultureInfo in a console application.

using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace AsposeCellsCulturePreview
{
    // A console‑based C# sample that asks for a locale (e.g., en‑US, fr‑FR), applies the corresponding CultureInfo to a Workbook, configures custom subtotal texts for a pivot table and custom chart legends via SettablePivotGlobalizationSettings and SettableChartGlobalizationSettings, displays the applied texts and number‑group separator, and saves the result as CulturePreview.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Prompt user for a culture name (e.g., "en-US", "fr-FR")
                Console.Write("Enter culture name (e.g., en-US, fr-FR): ");
                string cultureName = Console.ReadLine() ?? string.Empty;

                // Validate and create CultureInfo
                CultureInfo cultureInfo;
                try
                {
                    cultureInfo = new CultureInfo(cultureName);
                }
                catch (CultureNotFoundException)
                {
                    Console.WriteLine("Invalid culture. Falling back to invariant culture.");
                    cultureInfo = CultureInfo.InvariantCulture;
                }

                // Create a new workbook and apply the selected culture
                Workbook workbook = new Workbook();
                workbook.Settings.CultureInfo = cultureInfo;

                // -------------------------------------------------
                // Configure Pivot (subtotal) globalization settings
                // -------------------------------------------------
                var pivotSettings = new SettablePivotGlobalizationSettings();
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Custom Sum Total");
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "Custom Avg Total");

                // -------------------------------------------------
                // Configure Chart globalization settings
                // -------------------------------------------------
                var chartSettings = new SettableChartGlobalizationSettings();
                chartSettings.SetSeriesName("Custom Series");
                chartSettings.SetChartTitleName("Custom Chart Title");
                chartSettings.SetLegendIncreaseName("Increase");
                chartSettings.SetLegendDecreaseName("Decrease");
                chartSettings.SetOtherName("Other Category");

                // Combine both settings into a GlobalizationSettings instance
                var globalization = new GlobalizationSettings
                {
                    PivotSettings = pivotSettings,
                    ChartSettings = chartSettings
                };
                workbook.Settings.GlobalizationSettings = globalization;

                // -------------------------------------------------
                // Populate sample data
                // -------------------------------------------------
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("A");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("B");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("C");
                cells["B4"].PutValue(30);
                cells["A5"].PutValue("A");
                cells["B5"].PutValue(40);
                cells["A6"].PutValue("B");
                cells["B6"].PutValue(50);
                cells["A7"].PutValue("C");
                cells["B7"].PutValue(60);

                // -------------------------------------------------
                // Create a pivot table with subtotal
                // -------------------------------------------------
                int pivotIndex = sheet.PivotTables.Add("A1:B7", "D1", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data

                // Note: Subtotal settings are applied via GlobalizationSettings.
                // If explicit subtotal functions are required, use the appropriate API
                // available in the version of Aspose.Cells you are referencing.

                // -------------------------------------------------
                // Create a chart that uses the same data
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B7", true);
                chart.NSeries.CategoryData = "A2:A7";
                chart.Title.Text = "Demo Chart";

                // -------------------------------------------------
                // Preview the applied globalization texts
                // -------------------------------------------------
                Console.WriteLine();
                Console.WriteLine("=== Preview of Globalization Settings ===");
                Console.WriteLine($"Culture: {cultureInfo.Name}");
                Console.WriteLine($"Number group separator: '{cultureInfo.NumberFormat.NumberGroupSeparator}'");
                Console.WriteLine();

                Console.WriteLine("Pivot Subtotal Labels:");
                Console.WriteLine($" Sum: {pivotSettings.GetTextOfSubTotal(PivotFieldSubtotalType.Sum)}");
                Console.WriteLine($" Average: {pivotSettings.GetTextOfSubTotal(PivotFieldSubtotalType.Average)}");
                Console.WriteLine();

                Console.WriteLine("Chart Labels:");
                Console.WriteLine($" Series Name: {chartSettings.GetSeriesName()}");
                Console.WriteLine($" Chart Title: {chartSettings.GetChartTitleName()}");
                Console.WriteLine($" Legend Increase: {chartSettings.GetLegendIncreaseName()}");
                Console.WriteLine($" Legend Decrease: {chartSettings.GetLegendDecreaseName()}");
                Console.WriteLine($" Other: {chartSettings.GetOtherName()}");
                Console.WriteLine();

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "CulturePreview.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
