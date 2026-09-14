// Title: Check that a custom palette color remains after saving and reopening an Excel file with Aspose.Cells for .NET
// AI Prompts: Create a C# program using Aspose.Cells that changes a specific palette index to a custom RGB value, saves the workbook, reloads it, and verifies the color value matches the original. | Show how to add a slicer to a pivot table in Aspose.Cells and then confirm that the workbook’s custom palette entry persists after the file is saved and opened again.
// Common Searches: Aspose.Cells .NET verify custom palette entry after workbook save | C# test if Excel palette color persists when using ChangePalette | How to check color persistence in slicer‑linked workbook with Aspose.Cells | Validate IsColorInPalette after reloading an Excel file in C#
// Tags: change palette entry Aspose.Cells | custom RGB color persistence Excel .NET | slicer linked pivot table palette verification | IsColorInPalette after workbook reload | validate palette index after save

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The example creates a new workbook, modifies palette index 55 to a light orange RGB color, adds sample data, builds a pivot table, inserts a slicer, saves the file, reloads it, retrieves the color from the same palette index, compares it to the original custom color, and uses IsColorInPalette to confirm the color is still present, outputting the verification results.
class CustomSliceColorPersistenceDemo
{
    static void Main()
    {
        // Path for the temporary workbook file
        string filePath = "CustomSliceColorDemo.xlsx";

        // ---------- Create a new workbook ----------
        Workbook workbook = new Workbook();

        // Define a custom color and the palette index to modify (e.g., index 55)
        Color customColor = Color.FromArgb(255, 250, 223); // Light orange
        int paletteIndex = 55;

        // Change the palette entry to the custom color
        workbook.ChangePalette(customColor, paletteIndex);

        // Add sample data for a pivot table (required to create a slicer)
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Cells["A1"].PutValue("Product");
        dataSheet.Cells["A2"].PutValue("Apple");
        dataSheet.Cells["A3"].PutValue("Banana");
        dataSheet.Cells["A4"].PutValue("Orange");
        dataSheet.Cells["B1"].PutValue("Sales");
        dataSheet.Cells["B2"].PutValue(120);
        dataSheet.Cells["B3"].PutValue(80);
        dataSheet.Cells["B4"].PutValue(150);

        // Create a pivot table based on the data
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
        int pivotIdx = pivotSheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIdx];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

        // Add a slicer linked to the pivot table (the slicer itself does not store custom colors,
        // but its existence ensures the workbook contains slicer objects)
        Worksheet slicerSheet = workbook.Worksheets.Add("Slicer");
        int slicerIdx = slicerSheet.Slicers.Add(pivotTable, "A1", "Product");
        Slicer slicer = slicerSheet.Slicers[slicerIdx];
        slicer.StyleType = SlicerStyleType.SlicerStyleDark2; // any built‑in style

        // ---------- Save the workbook ----------
        workbook.Save(filePath);

        // ---------- Load the workbook ----------
        Workbook loadedWorkbook = new Workbook(filePath);

        // Retrieve the color from the palette after reload
        Color loadedColor = loadedWorkbook.Colors[paletteIndex];

        // Verify that the loaded color matches the custom color
        bool colorsMatch = loadedColor.ToArgb() == customColor.ToArgb();

        // Additionally, check that the color is recognized as being in the palette
        bool isInPalette = loadedWorkbook.IsColorInPalette(customColor);

        // Output verification results
        Console.WriteLine($"Custom color persisted: {colorsMatch}");
        Console.WriteLine($"Is custom color in palette after reload: {isInPalette}");
        Console.WriteLine($"Loaded palette color (ARGB): {loadedColor.ToArgb():X8}");
    }
}
