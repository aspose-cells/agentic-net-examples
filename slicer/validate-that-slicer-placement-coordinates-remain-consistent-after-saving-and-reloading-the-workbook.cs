// Title: Verify Slicer Position Persistence After Saving and Reloading with Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds a pivot table, inserts a slicer, sets explicit Shape.Left and Shape.Top values, saves the file, reloads it, and checks that the slicer's coordinates remain unchanged within a tiny tolerance.
// Keywords: Aspose.Cells slicer position | C# slicer coordinates persistence | Excel pivot table slicer location | save and reload workbook Aspose | .NET verify slicer placement | shape left top consistency | Excel automation regression test
// Common Searches: Aspose.Cells keep slicer location after save | C# check slicer left top after workbook reload | verify slicer placement persistence in Excel | Aspose.Cells .NET slicer shape coordinates | how to test slicer position stability
// Developer Intent: Confirm that a slicer's Left and Top properties are identical before and after the workbook is saved and opened again.
// Use Cases: Automated regression test to ensure dashboard slicers do not shift after file serialization. | Generate Excel reports with fixed slicer locations for consistent UI across deployments. | CI pipeline validation that slicer placement survives round‑trip save/load operations.
// AI Prompts: Write C# code using Aspose.Cells to add a slicer to a pivot table, set its Shape.Left and Shape.Top, save the workbook, reload it, and assert the coordinates are unchanged. | Create a reusable method that validates slicer position persistence after a workbook is saved and opened with Aspose.Cells for .NET. | Explain how to retrieve and compare slicer shape properties in a loaded workbook to verify placement stability.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, adds a pivot table, inserts a slicer, sets explicit Shape.Left and Shape.Top values, saves the file, reloads it, and checks that the slicer's coordinates remain unchanged within a tiny tolerance.
class SlicerPlacementValidation
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate sample data for the pivot table
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["A2"].PutValue("A");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["A3"].PutValue("B");
        ws.Cells["B3"].PutValue(20);
        ws.Cells["A4"].PutValue("A");
        ws.Cells["B4"].PutValue(30);

        // Add a pivot table based on the data range
        int pivotIdx = ws.PivotTables.Add("A1:B4", "D1", "Pivot1");
        PivotTable pt = ws.PivotTables[pivotIdx];
        pt.AddFieldToArea(PivotFieldType.Row, 0);   // Category field
        pt.AddFieldToArea(PivotFieldType.Data, 1);  // Value field
        pt.RefreshData();
        pt.CalculateData();

        // Add a slicer linked to the pivot table's first field (Category)
        int slicerIdx = ws.Slicers.Add(pt, "F1", "Category");
        Slicer slicer = ws.Slicers[slicerIdx];

        // Set explicit placement coordinates using the Shape object
        slicer.Shape.Left = 100;   // pixels from the left edge of the worksheet
        slicer.Shape.Top = 50;     // pixels from the top edge of the worksheet
        slicer.Shape.Width = 150;
        slicer.Shape.Height = 120;

        // Store the expected coordinates for later comparison
        double expectedLeft = slicer.Shape.Left;
        double expectedTop = slicer.Shape.Top;

        // Save the workbook to a file
        string filePath = "SlicerPlacementTest.xlsx";
        wb.Save(filePath);

        // Load the workbook back from the file
        Workbook loadedWb = new Workbook(filePath);
        Worksheet loadedWs = loadedWb.Worksheets[0];
        Slicer loadedSlicer = loadedWs.Slicers[slicerIdx];

        // Retrieve the actual coordinates after reload
        double actualLeft = loadedSlicer.Shape.Left;
        double actualTop = loadedSlicer.Shape.Top;

        // Validate that the coordinates are consistent
        bool leftMatches = Math.Abs(expectedLeft - actualLeft) < 0.001;
        bool topMatches = Math.Abs(expectedTop - actualTop) < 0.001;

        Console.WriteLine($"Left coordinate consistent: {leftMatches}");
        Console.WriteLine($"Top coordinate consistent: {topMatches}");
        Console.WriteLine($"Expected Left: {expectedLeft}, Actual Left: {actualLeft}");
        Console.WriteLine($"Expected Top: {expectedTop}, Actual Top: {actualTop}");
    }
}
