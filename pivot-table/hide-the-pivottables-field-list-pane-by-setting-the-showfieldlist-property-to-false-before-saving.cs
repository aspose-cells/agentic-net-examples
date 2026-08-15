// Title: Hide PivotTable Field List Pane with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add sample data, build a PivotTable on a separate sheet, set the PivotTable's EnableFieldList property to false, and save the file so the field‑list pane is not shown when the workbook is opened.
// Keywords: Aspose.Cells C# | Aspose.Cells .NET | PivotTable Hide Field List | EnableFieldList false | disable pivot field list pane | Excel automation Aspose | pivot UI hide | Aspose.Cells example | C# Excel PivotTable
// Common Searches: Aspose.Cells hide PivotTable field list C# | EnableFieldList property Aspose.Cells | disable pivot field list before saving workbook | C# code to hide PivotTable UI in Aspose.Cells | Aspose.Cells PivotTable field list pane example
// Developer Intent: The developer wants to prevent the PivotTable field‑list pane from appearing by disabling it before the workbook is saved.
// Use Cases: Generate a reporting workbook that opens with a clean PivotTable UI, preventing users from altering the layout. | Create a template where the PivotTable layout is locked, eliminating accidental changes to the field list. | Automate export of pivot data for distribution, ensuring the final Excel file shows no PivotTable UI elements.
// AI Prompts: Write C# code using Aspose.Cells to create a PivotTable and hide its field‑list pane before saving the workbook. | Explain the purpose of the EnableFieldList property in Aspose.Cells PivotTable and how to set it to false. | Show how to hide the field‑list pane for all PivotTables in a workbook with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, add sample data, build a PivotTable on a separate sheet, set the PivotTable's EnableFieldList property to false, and save the file so the field‑list pane is not shown when the workbook is opened.
class HidePivotFieldListDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["A4"].PutValue("A");
        sheet.Cells["B4"].PutValue(150);

        // Add a new worksheet that will contain the pivot table
        int pivotSheetIndex = workbook.Worksheets.Add(SheetType.Worksheet);
        Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
        pivotSheet.Name = "PivotTable";

        // Create the pivot table on the new sheet
        int pivotIndex = pivotSheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Row field: Category
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Data field: Value

        // Hide the field list pane for the pivot table
        pivotTable.EnableFieldList = false;

        // Save the workbook
        workbook.Save("HidePivotFieldListDemo.xlsx");
    }
}
