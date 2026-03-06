using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class CreateSlicerExample
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the existing workbook (lifecycle: load)
        Workbook workbook = new Workbook(inputPath);
        Worksheet sheet = workbook.Worksheets[0];

        // Define the data range that will be used for the pivot table (adjust as needed)
        // Here we assume the data is in columns A and B with a header row.
        string dataRange = "A1:B5";

        // Add a pivot table based on the data range (creation of pivot table)
        int pivotIndex = sheet.PivotTables.Add(dataRange, "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Configure the pivot table: first column as row field, second column as data field
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Column A
        pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Column B

        // Add a slicer linked to the pivot table.
        // The slicer will be placed with its upper‑left corner at row 20, column 12 (L20).
        // The field name "Fruit" must match a field in the pivot table's BaseFields.
        int slicerIndex = sheet.Slicers.Add(pivot, 20, 12, "Fruit");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Optional: set slicer properties (e.g., caption)
        slicer.Caption = "Fruit Slicer";

        // Save the modified workbook (lifecycle: save)
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}