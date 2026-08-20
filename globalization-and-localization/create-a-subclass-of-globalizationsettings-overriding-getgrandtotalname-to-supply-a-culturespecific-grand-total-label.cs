// Title: C# – Override GlobalizationSettings.GetGrandTotalName for Culture‑Specific Pivot Table Labels in Aspose.Cells
// Description: Shows how to subclass Aspose.Cells GlobalizationSettings, override GetGrandTotalName to return German labels for Sum, Average and Count, attach the custom settings to a workbook, build a pivot table, refresh it, and save the result as an Excel file.
// Keywords: Aspose.Cells | C# | GlobalizationSettings | GetGrandTotalName | pivot table localization | culture specific grand total | German label | custom globalization | Excel export | override method
// Common Searches: Aspose.Cells customize grand total text | C# override GetGrandTotalName | localize pivot table totals Aspose | German grand total label for pivot table | apply custom GlobalizationSettings workbook
// Developer Intent: Create a C# subclass of GlobalizationSettings that supplies localized grand‑total names and use it in a workbook’s pivot tables.
// Use Cases: Display German grand‑total captions (Sum, Average, Count) in automatically generated pivot tables. | Provide a fallback to the default label for consolidation functions that are not explicitly localized. | Reuse the same CustomGlobalizationSettings across multiple workbooks to ensure consistent language handling.
// AI Prompts: Generate a GlobalizationSettings subclass that returns French grand‑total names for Sum, Average, and Count and demonstrate its use in a pivot table. | Write code to attach a custom GlobalizationSettings instance to an existing workbook, refresh all pivot tables, and save the file. | Explain how to extend GetGrandTotalName to support additional functions such as Max, Min, and StdDev.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Subclass GlobalizationSettings to provide culture‑specific grand total names
// Shows how to subclass Aspose.Cells GlobalizationSettings, override GetGrandTotalName to return German labels for Sum, Average and Count, attach the custom settings to a workbook, build a pivot table, refresh it, and save the result as an Excel file.
public class CustomGlobalizationSettings : GlobalizationSettings
{
    // Override the method that returns the grand total label for a given function type
    public override string GetGrandTotalName(ConsolidationFunction functionType)
    {
        // Example: German culture specific labels
        return functionType switch
        {
            ConsolidationFunction.Sum => "Gesamtsumme",
            ConsolidationFunction.Average => "Gesamtdurchschnitt",
            ConsolidationFunction.Count => "Gesamtanzahl",
            _ => base.GetGrandTotalName(functionType) // Fallback to default for other functions
        };
    }
}

// Demo class that creates a workbook, applies the custom settings, and saves the file
public class GlobalizationGrandTotalDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for a pivot table
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Amount");
            worksheet.Cells["A2"].PutValue("Food");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["A3"].PutValue("Drink");
            worksheet.Cells["B3"].PutValue(80);
            worksheet.Cells["A4"].PutValue("Food");
            worksheet.Cells["B4"].PutValue(150);

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Create a pivot table to demonstrate the custom grand total label
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Row field: Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Data field: Amount
            pivotTable.DataFields[0].Function = ConsolidationFunction.Sum; // Use Sum function

            // Refresh and calculate the pivot table so that labels are generated
            pivotTable.RefreshData();   // Correct API call
            pivotTable.CalculateData();

            // Save the workbook (lifecycle save)
            workbook.Save("CustomGrandTotalDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during demo execution: {ex.Message}");
        }
    }
}

// Entry point for the application
public class Program
{
    public static void Main(string[] args)
    {
        GlobalizationGrandTotalDemo.Run();
    }
}
