// Title: Set slicer Shape.Placement to MoveAndSize and verify it stays unchanged after saving and reopening the workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate a new workbook, build a pivot table, attach a slicer, assign PlacementType.MoveAndSize to the slicer's Shape.Placement, save the file, reopen it, and read the placement value to ensure it remained MoveAndSize. | Modify an existing slicer in a .xlsx created with Aspose.Cells so that its shape uses the MoveAndSize placement mode, then persist the workbook and confirm the mode is retained after loading.
// Common Searches: Aspose.Cells C# keep slicer location after saving workbook | set slicer placement MoveAndSize in Aspose.Cells .NET | verify slicer shape placement persists after reloading Excel file using Aspose.Cells | C# Aspose.Cells slicer placement property resets on workbook reload | example of using Shape.Placement with slicers in Aspose.Cells
// Tags: Aspose.Cells slicer shape mode | persist slicer shape setting after workbook save | C# set slicer Shape.Placement property | pivot table slicer configuration Aspose.Cells | Excel file reload retains slicer settings .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the pivot, sets the slicer's Shape.Placement to PlacementType.MoveAndSize, saves the workbook as .xlsx, reloads it, and prints the placement to confirm the setting persisted.
class SlicerPlacementDemo
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

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Add a slicer linked to the pivot table
        int slicerIndex = worksheet.Slicers.Add(pivotTable, "E3", "Fruit");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Set the placement of the slicer using the non‑obsolete Shape.Placement property
        slicer.Shape.Placement = PlacementType.MoveAndSize;

        // Save the workbook
        string fileName = "SlicerPlacementDemo.xlsx";
        workbook.Save(fileName);

        // Reload the workbook to verify that the placement persists
        Workbook loadedWorkbook = new Workbook(fileName);
        Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
        Slicer loadedSlicer = loadedWorksheet.Slicers[0];

        // Output the placement after reload
        PlacementType placementAfterReload = loadedSlicer.Shape.Placement;
        Console.WriteLine("Slicer placement after reload: " + placementAfterReload);
    }
}
