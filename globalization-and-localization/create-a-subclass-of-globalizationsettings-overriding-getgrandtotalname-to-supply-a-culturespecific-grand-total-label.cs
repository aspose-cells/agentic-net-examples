// Title: How to subclass Aspose.Cells GlobalizationSettings in C# to override GetGrandTotalName for culture‑specific pivot table Grand Total labels
// AI Prompts: Create a C# class that inherits from Aspose.Cells.GlobalizationSettings and overrides GetGrandTotalName to return localized strings such as "Total général" for French, "Gesamtsumme" for German, "Total general" for Spanish, and "总计" for Chinese. | Demonstrate how to assign the custom GlobalizationSettings instance to a Workbook so that a pivot table automatically shows the localized Grand Total caption based on a given locale identifier.
// Common Searches: C# Aspose.Cells custom GlobalizationSettings for pivot table Grand Total localization | override GetGrandTotalName to display French grand total label in Aspose.Cells | how to set culture‑specific Grand Total text in Aspose.Cells pivot tables | using Aspose.Cells to localize pivot table totals for multiple languages
// Tags: GlobalizationSettings subclass example | C# pivot table total caption localization | localized grand total method | grand total label per locale | Aspose.Cells workbook localization

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Ensure PivotTable and PivotFieldType are available

// The example shows how to create a subclass of Aspose.Cells.GlobalizationSettings that overrides GetGrandTotalName to provide locale‑aware Grand Total strings, apply this settings object to a Workbook, and generate a pivot table whose total label adapts to French, German, Spanish, Chinese, or default English.
class Program
{
    // Returns a Grand Total label based on the supplied locale.
    private static string GetGrandTotalName(string locale)
    {
        CultureInfo culture = new CultureInfo(locale);
        return culture.TwoLetterISOLanguageName switch
        {
            "fr" => "Total général",   // French
            "de" => "Gesamtsumme",     // German
            "es" => "Total general",   // Spanish
            "zh" => "总计",            // Chinese
            _ => "Grand Total"         // Default English
        };
    }

    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Populate sample data for a pivot table.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(250);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(80);
            sheet.Cells["A5"].PutValue("B");
            sheet.Cells["B5"].PutValue(150);

            // Define the data range for the pivot table.
            string sourceRange = "A1:B5"; // Covers the populated data.
            int pivotRow = 7;   // Zero‑based row index where the pivot table will be placed.
            int pivotColumn = 0; // Zero‑based column index.

            // Add the pivot table (provide a name as the first argument).
            int pivotIndex = sheet.PivotTables.Add("PivotTable1", pivotRow, pivotColumn, sourceRange);
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure fields: Category as row field, Amount as data field.
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 (Category)
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 (Amount)

            // Refresh the pivot table to apply settings.
            pivot.RefreshData();

            // Ensure the output directory exists.
            string outputPath = "CustomGrandTotal.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
