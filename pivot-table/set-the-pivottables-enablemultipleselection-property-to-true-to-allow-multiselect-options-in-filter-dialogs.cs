// Title: Enable Multi‑Select Filtering on a PivotTable Page Field with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a PivotTable, place the "Category" field in the page (filter) area, and set its IsMultipleItemSelectionAllowed property to true so users can select multiple items in the filter dialog. The workbook is saved as an .xlsx file.
// Keywords: Aspose.Cells | .NET | C# | PivotTable | EnableMultipleSelection | IsMultipleItemSelectionAllowed | multi select filter | page field | pivot filter | Excel sample code
// Common Searches: Aspose.Cells enable multi select pivot filter C# | Set IsMultipleItemSelectionAllowed true Aspose.Cells | Allow multiple items in PivotTable page field Aspose | C# code for multi‑select pivot table filter Aspose.Cells | EnableMultipleSelection property Aspose.Cells PivotTable
// Developer Intent: Activate multi‑selection for a PivotTable page (filter) field using Aspose.Cells for .NET.
// Use Cases: Sales dashboard where analysts can filter by several product categories simultaneously. | Regional performance report that lets users pick multiple territories in a pivot filter. | Financial model requiring combined selection of multiple expense types in an Excel pivot table.
// AI Prompts: Show C# code to set IsMultipleItemSelectionAllowed = true for a PivotTable page field with Aspose.Cells. | Provide an Aspose.Cells example that creates a PivotTable with multi‑select filter options. | Explain how to enable multiple item selection for all page fields in an Aspose.Cells PivotTable.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, add a PivotTable, place the "Category" field in the page (filter) area, and set its IsMultipleItemSelectionAllowed property to true so users can select multiple items in the filter dialog. The workbook is saved as an .xlsx file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "Amount";
        sheet.Cells["A2"].Value = "Food";
        sheet.Cells["B2"].Value = 100;
        sheet.Cells["A3"].Value = "Drink";
        sheet.Cells["B3"].Value = 150;
        sheet.Cells["A4"].Value = "Food";
        sheet.Cells["B4"].Value = 200;

        // Add a pivot table to the worksheet
        // Source range: A1:B4, Destination: D1, Name: PivotTable1
        int ptIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[ptIndex];

        // Add fields: make "Category" a page (filter) field and "Amount" a data field
        pivotTable.AddFieldToArea(PivotFieldType.Page, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Enable multiple item selection for the page field
        // This allows users to select multiple items in the filter dialog
        PivotField pageField = pivotTable.PageFields[0];
        pageField.IsMultipleItemSelectionAllowed = true;

        // Save the workbook to a file
        workbook.Save("PivotTable_MultiSelect.xlsx");
    }
}
