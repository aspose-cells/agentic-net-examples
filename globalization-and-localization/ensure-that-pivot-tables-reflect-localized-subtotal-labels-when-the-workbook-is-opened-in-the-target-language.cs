// Title: Aspose.Cells .NET – Localize Pivot Table Subtotal Captions
// Description: Demonstrates how to use SettablePivotGlobalizationSettings in Aspose.Cells for .NET to replace default pivot table subtotal texts (Sum, Count, Average, Max, Min) with custom strings for any target language, refresh the pivot, and save the workbook.
// Keywords: Aspose.Cells | pivot table localization | custom subtotal captions | SettablePivotGlobalizationSettings | C# Excel i18n | globalization settings | Excel multi‑language reports | .NET Excel export | internationalization pivot table | localized Excel subtotals
// Common Searches: change pivot subtotal text Aspose.Cells | localize Excel pivot table labels .NET | SettablePivotGlobalizationSettings example | customize sum count average captions in Excel | Aspose.Cells pivot table language support
// Developer Intent: Show pivot tables with subtotal labels translated to the workbook’s target language.
// Use Cases: Create Excel dashboards for global audiences where subtotal rows appear in the local language. | Prepare a reusable workbook template that automatically applies region‑specific subtotal terminology. | Meet compliance or branding guidelines by displaying translated sum, count, and average labels in exported reports.
// AI Prompts: Generate C# code that sets French subtitles for pivot table subtotals using Aspose.Cells. | Explain how to read existing subtotal captions from a pivot table and replace them at runtime. | Provide a step‑by‑step guide to apply different localization settings for multiple languages in one workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsPivotLocalizationDemo
{
    // Demonstrates how to use SettablePivotGlobalizationSettings in Aspose.Cells for .NET to replace default pivot table subtotal texts (Sum, Count, Average, Max, Min) with custom strings for any target language, refresh the pivot, and save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("A");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("B");
            dataSheet.Cells["B4"].PutValue(30);
            dataSheet.Cells["A5"].PutValue("B");
            dataSheet.Cells["B5"].PutValue(40);

            // Add a pivot table based on the data range
            int pivotIndex = dataSheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

            // Create a SettablePivotGlobalizationSettings instance to customize subtotal texts
            SettablePivotGlobalizationSettings localizationSettings = new SettablePivotGlobalizationSettings();

            // Set custom localized texts for various subtotal types
            localizationSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Σ Total");          // Sum
            localizationSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "Count Total");   // Count
            localizationSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "Avg Total");   // Average
            localizationSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Max, "Maximum Total");   // Max
            localizationSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Min, "Minimum Total");   // Min

            // Assign the customized settings to the workbook's globalization settings
            workbook.Settings.GlobalizationSettings.PivotSettings = localizationSettings;

            // Refresh and calculate the pivot table so that the new labels take effect
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("LocalizedPivotSubtotals.xlsx");
        }
    }
}
