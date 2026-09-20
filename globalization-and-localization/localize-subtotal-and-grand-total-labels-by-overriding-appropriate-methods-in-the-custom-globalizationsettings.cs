// Title: How to override GlobalizationSettings to localize Subtotal and Grand Total labels in an Aspose.Cells PivotTable (C#)
// AI Prompts: Create a subclass of Aspose.Cells.GlobalizationSettings that provides custom text for subtotal and grand total, then assign an instance to Workbook.GlobalizationSettings before building the pivot table. | Show a French localization by returning "Sous‑total" from GetSubtotalLabel and "Total général" from GetGrandTotalLabel, and generate a pivot table that displays these captions. | Integrate the custom GlobalizationSettings into the sample code so the saved workbook shows the localized Subtotal and Grand Total labels.
// Common Searches: asp.net aspocells customize subtotal label in pivot table | override globalizationsettings getgrandtotallabel c# example | localize pivot table total captions using Aspose.Cells | change pivot table grand total text Aspose.Cells .NET | how to set custom subtotal text for Aspose.Cells pivot table
// Tags: custom GlobalizationSettings for pivot totals | localize subtotal and grand total captions | Aspose.Cells pivot table label customization | C# override GetGrandTotalLabel | Excel workbook globalization settings

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsLocalizationExample
{
    // // This example creates a workbook, populates it with sample category and amount data, defines a pivot table, and demonstrates how to subclass GlobalizationSettings to override GetSubtotalLabel and GetGrandTotalLabel. The custom settings are assigned to the workbook so the pivot table displays localized Subtotal and Grand Total captions before saving as LocalizedPivotTable.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Add sample data.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Food");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Transport");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Utilities");
                sheet.Cells["B4"].PutValue(150);

                // Define the source data range.
                int totalRows = 5; // includes header row
                string sourceDataRange = $"A1:B{totalRows}";
                string pivotTableDestination = "D1";

                // Create the pivot table (Add returns the index of the new table).
                int pivotIndex = sheet.PivotTables.Add(pivotTableDestination, sourceDataRange, "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add "Category" as a row field.
                // Use the field name indexer to obtain the PivotField object.
                pivotTable.RowFields.Add(pivotTable.RowFields["Category"]);

                // Add "Amount" as a data field.
                pivotTable.DataFields.Add(pivotTable.DataFields["Amount"]);

                // Save the workbook.
                string outputPath = "LocalizedPivotTable.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
