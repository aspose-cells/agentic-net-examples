// Title: Check that a slicer’s Left and Top coordinates stay the same after saving and reloading an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# program that adds a slicer to a pivot table, sets its Shape.Left and Shape.Top values, saves the workbook, reloads it, and asserts the coordinates are unchanged within a small tolerance. | Adapt an existing Aspose.Cells slicer example to output a pass/fail result based on whether the saved and loaded slicer positions match.
// Common Searches: aspnet cells keep slicer position after workbook save | c# verify slicer shape left top values after reopening Excel file | aspose.cells slicer placement persistence across save and load | how to test slicer coordinates consistency in saved workbook using Aspose.Cells
// Tags: slicer shape position persistence Aspose.Cells | validate slicer coordinates after workbook reload | Aspose.Cells set slicer left top | pivot table slicer placement verification | Excel file save reload slicer shape

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

// The code creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the pivot, sets the slicer's Left and Top coordinates, saves the file, reloads it, and confirms that the slicer's position remains unchanged.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Fruit");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["B4"].PutValue(200);

        // Add a pivot table
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Add a slicer linked to the pivot table
        int slicerIndex = worksheet.Slicers.Add(pivotTable, "F1", "Fruit");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Set placement coordinates using the Shape object
        slicer.Shape.Left = 50; // pixels from the left edge of the worksheet
        slicer.Shape.Top = 30;  // pixels from the top edge of the worksheet

        // Store original coordinates for later comparison
        double originalLeft = slicer.Shape.Left;
        double originalTop = slicer.Shape.Top;

        // Save the workbook
        string filePath = "SlicerPlacementTest.xlsx";
        workbook.Save(filePath);

        // Load the workbook back
        Workbook loadedWorkbook = new Workbook(filePath);
        Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
        Slicer loadedSlicer = loadedWorksheet.Slicers[slicerIndex];

        // Retrieve coordinates after reload
        double loadedLeft = loadedSlicer.Shape.Left;
        double loadedTop = loadedSlicer.Shape.Top;

        // Validate that the coordinates remain consistent
        bool leftConsistent = Math.Abs(originalLeft - loadedLeft) < 0.001;
        bool topConsistent = Math.Abs(originalTop - loadedTop) < 0.001;

        Console.WriteLine($"Left coordinate consistent: {leftConsistent}");
        Console.WriteLine($"Top coordinate consistent: {topConsistent}");
    }
}
