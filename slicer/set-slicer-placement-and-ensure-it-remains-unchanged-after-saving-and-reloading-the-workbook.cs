// Title: Persist Slicer Placement (MoveAndSize) After Save with Aspose.Cells for .NET
// Description: Demonstrates how to set a slicer's Placement to MoveAndSize, save the workbook, reload it, and verify that the placement setting is retained using Aspose.Cells for C#.
// Keywords: Aspose.Cells slicer placement | PlacementType.MoveAndSize C# | slicer persistence after save | Aspose.Cells pivot table slicer example | C# Excel slicer placement | retain slicer settings Aspose | Excel slicer MoveAndSize Aspose.Cells | slicer placement property .NET
// Common Searches: how to set slicer placement move and size with Aspose.Cells | does slicer placement persist after saving workbook in C# | Aspose.Cells example for slicer placement property | C# code to verify slicer placement after reload | Aspose.Cells keep slicer settings after file save
// Developer Intent: Set a slicer's Placement to MoveAndSize and confirm the setting survives workbook save and reload.
// Use Cases: Create a pivot table, add a linked slicer, and make the slicer move and resize with its cells. | Save the Excel file and reload it to ensure the slicer placement remains unchanged. | Log or display the slicer Placement value after reloading to validate persistence.
// AI Prompts: Generate C# code with Aspose.Cells that sets a slicer's Placement to MoveAndSize, saves the workbook, reloads it, and prints the placement value. | Explain the behavior of PlacementType.MoveAndSize for slicers in Excel files created with Aspose.Cells. | Write a method that iterates over all slicers in a worksheet and applies PlacementType.MoveAndSize before saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;
using Aspose.Cells.Drawing;

// Demonstrates how to set a slicer's Placement to MoveAndSize, save the workbook, reload it, and verify that the placement setting is retained using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].Value = "Fruit";
        worksheet.Cells["A2"].Value = "Apple";
        worksheet.Cells["A3"].Value = "Orange";
        worksheet.Cells["A4"].Value = "Banana";
        worksheet.Cells["B1"].Value = "Sales";
        worksheet.Cells["B2"].Value = 100;
        worksheet.Cells["B3"].Value = 150;
        worksheet.Cells["B4"].Value = 200;

        // Add a pivot table based on the data
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Add a slicer linked to the pivot table
        int slicerIndex = worksheet.Slicers.Add(pivotTable, "F1", "Fruit");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Set the slicer's placement so it moves and sizes with cells
        slicer.Placement = PlacementType.MoveAndSize;

        // Save the workbook
        string fileName = "SlicerPlacementDemo.xlsx";
        workbook.Save(fileName);

        // Reload the workbook to verify the placement persists
        Workbook reloadedWorkbook = new Workbook(fileName);
        Worksheet reloadedWorksheet = reloadedWorkbook.Worksheets[0];
        Slicer reloadedSlicer = reloadedWorksheet.Slicers[0];

        // Output the placement after reload
        Console.WriteLine("Slicer placement after reload: " + reloadedSlicer.Placement);
    }
}
