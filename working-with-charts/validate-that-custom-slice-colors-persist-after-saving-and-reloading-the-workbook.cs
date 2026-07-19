// Title: C# – Verify Slicer Style Color Persistence After Saving with Aspose.Cells
// Description: This example creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the pivot, applies the built‑in SlicerStyleDark2, saves the file, reloads it, and checks that the slicer's StyleType (and its colors) remain unchanged.
// Keywords: Aspose.Cells | C# | slicer style persistence | Excel slicer color save | pivot table slicer | SlicerStyleDark2 | workbook reload verification | unit test | Excel automation | style retention
// Common Searches: Aspose.Cells verify slicer style after save | C# check slicer color persistence in Excel | how to keep slicer formatting when reloading workbook | test slicer style retention with Aspose.Cells | persist Excel slicer colors .NET
// Developer Intent: Confirm that a slicer's visual style, including its color scheme, is preserved when an Excel workbook is saved and later reopened using Aspose.Cells for .NET.
// Use Cases: Automated regression test to ensure slicer formatting does not change across library versions. | Generating reports where slicer appearance must stay consistent for end users. | Validating that custom or built‑in slicer styles survive the save‑load cycle in CI pipelines.
// AI Prompts: Generate a NUnit test that asserts the slicer StyleType before and after saving the workbook with Aspose.Cells. | Show code to compare all visual properties of a slicer (style, caption, size) after reloading a workbook. | Provide an example that applies a custom color scheme to a slicer and verifies the colors persist after the file is saved and reopened.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// This example creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the pivot, applies the built‑in SlicerStyleDark2, saves the file, reloads it, and checks that the slicer's StyleType (and its colors) remain unchanged.
class ValidateSlicerColorPersistence
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook wb = new Workbook();
        Worksheet dataSheet = wb.Worksheets[0];
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["A2"].PutValue("Apple");
        dataSheet.Cells["A3"].PutValue("Banana");
        dataSheet.Cells["A4"].PutValue("Orange");
        dataSheet.Cells["B1"].PutValue("Sales");
        dataSheet.Cells["B2"].PutValue(100);
        dataSheet.Cells["B3"].PutValue(150);
        dataSheet.Cells["B4"].PutValue(200);

        // Create a pivot table based on the data
        Worksheet pivotSheet = wb.Worksheets.Add("Pivot");
        int pivotIdx = pivotSheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
        PivotTable pt = pivotSheet.PivotTables[pivotIdx];
        pt.AddFieldToArea(PivotFieldType.Row, 0);
        pt.AddFieldToArea(PivotFieldType.Data, 1);

        // Add a slicer linked to the pivot table
        Worksheet slicerSheet = wb.Worksheets.Add("Slicer");
        int slicerIdx = slicerSheet.Slicers.Add(pt, "A1", "Category");
        Slicer slicer = slicerSheet.Slicers[slicerIdx];

        // Set a built‑in slicer style (controls the button colors)
        slicer.StyleType = SlicerStyleType.SlicerStyleDark2;

        // Save the workbook
        string filePath = "SlicerColorPersistence.xlsx";
        wb.Save(filePath);

        // Reload the workbook
        Workbook loadedWb = new Workbook(filePath);
        Slicer loadedSlicer = loadedWb.Worksheets["Slicer"].Slicers[0];

        // Verify that the slicer style (and thus its colors) persisted after reload
        Console.WriteLine("Original Style: " + slicer.StyleType);
        Console.WriteLine("Loaded Style : " + loadedSlicer.StyleType);
        Console.WriteLine("Colors persisted: " + (slicer.StyleType == loadedSlicer.StyleType));
    }
}
