using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class CreateSlicerForPivot
{
    static void Main()
    {
        // Paths to the input and output Excel files
        string inputPath = "input.xlsx";
        string outputPath = "output_with_slicer.xlsx";

        // Load the existing workbook that already contains a pivot table
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure the worksheet has at least one pivot table
        if (sheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        // Retrieve the first pivot table
        PivotTable pivot = sheet.PivotTables[0];

        // Determine a base field name to use for the slicer.
        // Here we use the first field from the pivot's BaseFields collection.
        if (pivot.BaseFields.Count == 0)
        {
            Console.WriteLine("Pivot table has no base fields to attach a slicer.");
            return;
        }
        string baseFieldName = pivot.BaseFields[0].Name;

        // Add a slicer linked to the pivot table.
        // The slicer will be placed with its upper‑left corner at cell E3.
        int slicerIndex = sheet.Slicers.Add(pivot, "E3", baseFieldName);

        // Optionally retrieve the slicer object for further customization
        Slicer slicer = sheet.Slicers[slicerIndex];
        Console.WriteLine($"Slicer '{slicer.Name}' added at cell E3 for field '{baseFieldName}'.");

        // Save the modified workbook
        workbook.Save(outputPath, SaveFormat.Xlsx);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}