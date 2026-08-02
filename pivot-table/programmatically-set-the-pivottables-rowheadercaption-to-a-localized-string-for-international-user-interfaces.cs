// Title: Aspose.Cells .NET – Localize PivotTable Row Header Caption and Row Labels with SettablePivotGlobalizationSettings
// Description: Demonstrates how to create a workbook, add sample data, build a PivotTable, set a custom RowHeaderCaption, and replace the built‑in "Row Labels" text with a localized string using SettablePivotGlobalizationSettings before refreshing and saving the file.
// Keywords: Aspose.Cells | C# PivotTable localization | RowHeaderCaption | SettablePivotGlobalizationSettings | globalization settings Excel | .NET Excel PivotTable | localize Row Labels | international Excel UI | Aspose.Cells PivotTable example
// Common Searches: Aspose.Cells change Row Header Caption C# | localize PivotTable Row Labels Aspose.Cells | SettablePivotGlobalizationSettings example .NET | how to rename Row Labels in Aspose.Cells PivotTable | globalize Excel PivotTable text with Aspose
// Developer Intent: Apply a language‑specific caption to a PivotTable row header and override the default "Row Labels" label programmatically.
// Use Cases: Generate an Excel workbook with a PivotTable that displays row headers in the target language. | Replace the generic "Row Labels" label with a custom localized string for international user interfaces. | Refresh and calculate the PivotTable after applying globalization settings to ensure the new captions appear in the saved file.
// AI Prompts: Provide C# code using Aspose.Cells to set a localized RowHeaderCaption and change the built‑in "Row Labels" text via SettablePivotGlobalizationSettings. | Explain step‑by‑step how to apply globalization settings to an Aspose.Cells workbook so that all PivotTable row labels are translated. | Show an example that creates sample data, builds a PivotTable, customizes its row header caption, refreshes the table, and saves the workbook with localized labels.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;   // For SettableGlobalizationSettings

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, build a PivotTable, set a custom RowHeaderCaption, and replace the built‑in "Row Labels" text with a localized string using SettablePivotGlobalizationSettings before refreshing and saving the file.
    public class PivotTableRowHeaderLocalizationDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet (data source)
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate sample data for the pivot table
                var cells = dataSheet.Cells;
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Amount");
                cells["A2"].PutValue("Food");
                cells["B2"].PutValue(120);
                cells["A3"].PutValue("Beverage");
                cells["B3"].PutValue(80);
                cells["A4"].PutValue("Food");
                cells["B4"].PutValue(150);
                cells["A5"].PutValue("Beverage");
                cells["B5"].PutValue(70);

                // Add a new worksheet to host the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Create the pivot table (source range, destination cell, name)
                int pivotIndex = pivotSheet.PivotTables.Add(
                    "=Data!A1:B5",   // source range (using sheet name)
                    "A3",            // destination cell in pivot sheet
                    "SalesPivot");   // pivot table name

                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Set a custom caption for the row header directly
                pivotTable.RowHeaderCaption = "Localized Row Header";

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // ------------------------------------------------------------
                // Globalization settings: change the text of the built‑in
                // "Row Labels" label to a localized string.
                // ------------------------------------------------------------
                // Create a SettablePivotGlobalizationSettings instance
                SettablePivotGlobalizationSettings pivotGlobalSettings = new SettablePivotGlobalizationSettings();
                // Set the custom text for "Row Labels"
                pivotGlobalSettings.SetTextOfRowLabels("Localized Row Labels");

                // Wrap the pivot settings into SettableGlobalizationSettings
                SettableGlobalizationSettings globalizationSettings = new SettableGlobalizationSettings
                {
                    PivotSettings = pivotGlobalSettings
                };

                // Apply the globalization settings to the workbook
                workbook.Settings.GlobalizationSettings = globalizationSettings;

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PivotTableRowHeaderLocalization.xlsx");
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableRowHeaderLocalizationDemo.Run();
        }
    }
}
