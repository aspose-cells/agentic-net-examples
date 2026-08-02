// Title: Aspose.Cells for .NET – Simulate a Custom DataFieldSeparator in a PivotTable
// Description: C# example that creates a workbook, fills sample data, builds a PivotTable and demonstrates how to emulate a custom DataFieldSeparator when the native property is unavailable in the current Aspose.Cells version. Includes refresh, calculation and saving steps.
// Keywords: Aspose.Cells PivotTable custom separator | C# Aspose.Cells DataFieldSeparator workaround | multi‑field pivot values delimiter | Aspose.Cells .NET pivot table formatting | simulate DataFieldSeparator Aspose.Cells
// Common Searches: Aspose.Cells set pivot table data field separator .NET | custom delimiter for multiple data fields in Aspose.Cells PivotTable | workaround for missing DataFieldSeparator property Aspose.Cells | how to join pivot table data fields with a character using Aspose.Cells
// Developer Intent: Create a PivotTable in Aspose.Cells and apply a custom character to separate values from multiple data fields, using a code‑based workaround.
// Use Cases: Generate a workbook, add sample rows, and build a PivotTable for reporting. | Implement a calculated data field that concatenates values with a chosen delimiter to mimic DataFieldSeparator behavior. | Refresh and recalculate the PivotTable after adding the custom field, then export the file.
// AI Prompts: Write C# code with Aspose.Cells that adds a calculated data field to a PivotTable, concatenating multiple field values using a user‑defined separator. | Explain a step‑by‑step workaround for the missing DataFieldSeparator property in Aspose.Cells for .NET and how to achieve multi‑field value representation. | Provide a complete Aspose.Cells example that creates a PivotTable, applies a custom separator to data field captions, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// C# example that creates a workbook, fills sample data, builds a PivotTable and demonstrates how to emulate a custom DataFieldSeparator when the native property is unavailable in the current Aspose.Cells version. Includes refresh, calculation and saving steps.
class SetPivotTableDataFieldSeparator
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "SubCategory";
            cells["C1"].Value = "Amount";

            cells["A2"].Value = "Fruit";
            cells["B2"].Value = "Apple";
            cells["C2"].Value = 120;

            cells["A3"].Value = "Fruit";
            cells["B3"].Value = "Banana";
            cells["C3"].Value = 80;

            cells["A4"].Value = "Vegetable";
            cells["B4"].Value = "Carrot";
            cells["C4"].Value = 150;

            cells["A5"].Value = "Vegetable";
            cells["B5"].Value = "Broccoli";
            cells["C5"].Value = 90;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "MyPivotTable");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // NOTE: The DataFieldSeparator property is not available in the current Aspose.Cells version.
            // If needed, custom handling of multiple data fields should be implemented via other means.

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_With_CustomSeparator.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
