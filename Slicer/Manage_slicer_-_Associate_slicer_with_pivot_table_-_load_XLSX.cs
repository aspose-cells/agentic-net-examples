using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load the existing workbook (lifecycle rule: load)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet that contains the pivot table
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one pivot table
        if (sheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        // Get the first pivot table
        PivotTable pivot = sheet.PivotTables[0];

        // Ensure the pivot table has base fields to bind the slicer to
        if (pivot.BaseFields.Count == 0)
        {
            Console.WriteLine("Pivot table has no base fields.");
            return;
        }

        // Use the name of the first base field for the slicer
        string baseFieldName = pivot.BaseFields[0].Name;

        // Add a slicer linked to the pivot table (rule: SlicerCollection.Add with dest cell and field name)
        int slicerIndex = sheet.Slicers.Add(pivot, "E1", baseFieldName);
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Refresh the slicer to reflect current data
        slicer.Refresh();

        // Save the modified workbook (lifecycle rule: save)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}