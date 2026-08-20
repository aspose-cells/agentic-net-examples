// Title: C# – Subclass GlobalizationSettings to Localize Pivot Table Total Labels in Aspose.Cells
// Description: Demonstrates how to inherit from Aspose.Cells.GlobalizationSettings, override GetTotalName to return custom strings for Sum, Average and Count, assign the subclass to a Workbook, build a pivot table, and save the result with localized total labels.
// Keywords: Aspose.Cells | C# | GlobalizationSettings | GetTotalName | pivot table localization | custom subtotal label | Excel total name override | ConsolidationFunction Sum | ConsolidationFunction Average | ConsolidationFunction Count | localized Excel reports
// Common Searches: override GetTotalName Aspose.Cells | custom GlobalizationSettings C# example | localize pivot table total names | change subtotal label in Aspose.Cells | Aspose.Cells pivot table localization tutorial
// Developer Intent: Create a subclass of GlobalizationSettings that overrides GetTotalName to supply localized total labels for pivot tables.
// Use Cases: Display pivot table totals in the end‑user's language without modifying the source data. | Apply consistent branding by using company‑specific terminology for Sum, Average, and Count totals. | Extend the globalization layer to support additional consolidation functions for financial or statistical reports.
// AI Prompts: Write C# code that defines a CustomGlobalizationSettings class overriding GetTotalName for Sum, Average, and Count and applies it to a workbook with a pivot table. | Explain how to add custom total names for other ConsolidationFunction values in Aspose.Cells. | Show how to programmatically verify that the localized total labels appear in the generated Excel file after refreshing the pivot table.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to inherit from Aspose.Cells.GlobalizationSettings, override GetTotalName to return custom strings for Sum, Average and Count, assign the subclass to a Workbook, build a pivot table, and save the result with localized total labels.
public class CustomGlobalizationSettings : GlobalizationSettings
{
    // Override GetTotalName to return localized labels for different functions
    public override string GetTotalName(ConsolidationFunction functionType)
    {
        return functionType switch
        {
            ConsolidationFunction.Sum => "Localized Sum",
            ConsolidationFunction.Average => "Localized Average",
            ConsolidationFunction.Count => "Localized Count",
            _ => base.GetTotalName(functionType)
        };
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Apply the custom globalization settings
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Populate sample data
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Value");
        cells["A2"].PutValue("A");
        cells["B2"].PutValue(10);
        cells["A3"].PutValue("B");
        cells["B3"].PutValue(20);
        cells["A4"].PutValue("C");
        cells["B4"].PutValue(30);

        // Create a pivot table to trigger the total name usage
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, 0); // Row field
        int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, 1); // Data field
        pivot.DataFields[dataFieldIdx].Function = ConsolidationFunction.Sum; // Use Sum function

        // Refresh and calculate the pivot table
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook
        workbook.Save("CustomGlobalizationSettings.xlsx");
    }
}
