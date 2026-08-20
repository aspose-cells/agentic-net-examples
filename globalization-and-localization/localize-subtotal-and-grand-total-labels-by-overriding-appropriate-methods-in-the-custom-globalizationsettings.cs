// Title: Localize Pivot Table Subtotal & Grand Total Labels with Custom GlobalizationSettings (C# Aspose.Cells)
// Description: Demonstrates how to subclass SettablePivotGlobalizationSettings to replace English subtotal and grand total texts with Chinese equivalents, apply the settings to a workbook via reflection for version‑agnostic compatibility, build a simple pivot table, and save the result as LocalizedPivot.xlsx.
// Keywords: Aspose.Cells pivot localization | C# custom GlobalizationSettings | override subtotal label | override grand total label | Chinese pivot table totals | SettablePivotGlobalizationSettings example | reflection workbook settings | pivot table total text customization
// Common Searches: Aspose.Cells change pivot subtotal text | C# set custom grand total label in pivot table | localize Aspose.Cells pivot totals to Chinese | SettablePivotGlobalizationSettings usage | how to apply custom globalization settings with reflection
// Developer Intent: Replace the default English subtotal and grand total captions in a pivot table with user‑defined, localized strings.
// Use Cases: Generate reports for Chinese audiences where pivot totals read “合计”, “平均值”, etc. | Maintain a single code base that works across multiple Aspose.Cells versions by using reflection to set PivotGlobalizationSettings. | Create reusable globalization classes for enterprise‑wide spreadsheet localization.
// AI Prompts: Write a C# class that inherits SettablePivotGlobalizationSettings and returns Chinese text for each PivotFieldSubtotalType and for the grand total. | Show how to assign a custom GlobalizationSettings instance to a Workbook using reflection, with fallback for older Aspose.Cells releases. | Provide a complete example that builds a pivot table, applies the custom settings, refreshes the data, and saves the workbook with localized total labels.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to subclass SettablePivotGlobalizationSettings to replace English subtotal and grand total texts with Chinese equivalents, apply the settings to a workbook via reflection for version‑agnostic compatibility, build a simple pivot table, and save the result as LocalizedPivot.xlsx.
class CustomGlobalizationSettings : SettablePivotGlobalizationSettings
{
    // Override subtotal text for each subtotal type
    public override string GetTextOfSubTotal(PivotFieldSubtotalType subTotalType)
    {
        return subTotalType switch
        {
            PivotFieldSubtotalType.Sum => "合计",          // Sum
            PivotFieldSubtotalType.Average => "平均值",   // Average
            PivotFieldSubtotalType.Count => "计数",      // Count
            PivotFieldSubtotalType.Max => "最大值",       // Max
            PivotFieldSubtotalType.Min => "最小值",       // Min
            _ => base.GetTextOfSubTotal(subTotalType)
        };
    }

    // Override grand total label
    public override string GetTextOfGrandTotal()
    {
        return "总计"; // Grand Total
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("A");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("B");
            sheet.Cells["B4"].PutValue(30);

            // Add a pivot table
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

            // Apply custom globalization settings for pivot tables (using reflection for compatibility)
            try
            {
                var settingsObj = workbook.Settings;
                var prop = settingsObj.GetType().GetProperty("PivotGlobalizationSettings");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(settingsObj, new CustomGlobalizationSettings());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to set PivotGlobalizationSettings. {ex.Message}");
            }

            // Refresh pivot data and calculate
            pivot.RefreshData();      // Refreshes the pivot cache
            pivot.CalculateData();    // Calculates pivot values

            // Save the workbook
            workbook.Save("LocalizedPivot.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
