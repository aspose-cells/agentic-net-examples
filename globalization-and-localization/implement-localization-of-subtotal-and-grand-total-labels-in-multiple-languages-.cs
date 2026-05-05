using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    // Demonstrates how to localize PivotTable subtotal and grand total labels
    // for different languages in a workbook loaded from an existing XLSX file.
    class PivotLocalizationDemo
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create globalization settings that allow modifications
            SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings();

            // Create pivot‑specific settings
            SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();

            // Apply English (default) localization
            ApplyEnglishLocalization(pivotSettings);

            // Uncomment one of the following lines to switch language
            //ApplyFrenchLocalization(pivotSettings);
            //ApplyGermanLocalization(pivotSettings);

            // Attach the pivot settings to the global settings
            globalSettings.PivotSettings = pivotSettings;

            // Assign the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalSettings;

            // Refresh all pivot tables so that the new labels take effect
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pt in sheet.PivotTables)
                {
                    pt.RefreshData();
                    pt.CalculateData();
                }
            }

            // Save the modified workbook (replace with desired output path)
            workbook.Save("output.xlsx");
        }

        // Sets English labels (default)
        private static void ApplyEnglishLocalization(SettablePivotGlobalizationSettings settings)
        {
            // Total and Grand Total
            settings.SetTextOfTotal("Total");
            settings.SetTextOfGrandTotal("Grand Total");

            // Subtotal types
            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Sum");
            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "Count");
            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "Average");
            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Max, "Maximum");
            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Min, "Minimum");
        }

        // Sets French labels
        private static void ApplyFrenchLocalization(SettablePivotGlobalizationSettings settings)
        {
            settings.SetTextOfTotal("Total");
            settings.SetTextOfGrandTotal("Total général");

            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Somme");
            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "Nombre");
            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "Moyenne");
            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Max, "Maximum");
            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Min, "Minimum");
        }

        // Sets German labels
        private static void ApplyGermanLocalization(SettablePivotGlobalizationSettings settings)
        {
            settings.SetTextOfTotal("Summe");
            settings.SetTextOfGrandTotal("Gesamtsumme");

            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Summe");
            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "Anzahl");
            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "Durchschnitt");
            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Max, "Maximum");
            settings.SetTextOfSubTotal(PivotFieldSubtotalType.Min, "Minimum");
        }
    }
}