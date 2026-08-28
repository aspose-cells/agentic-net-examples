// Title: Add a slicer linked to the first base field of a pivot table in an existing Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file, finds the first pivot table, adds a slicer for its first base field at cell E3, sets a custom caption, and saves the workbook using Aspose.Cells. | Create a method that iterates over all base fields of a given pivot table and inserts a separate slicer for each field, positioning them sequentially on the worksheet with Aspose.Cells. | Modify the workbook saving routine to keep existing slicer settings intact while appending new slicers programmatically in Aspose.Cells.
// Common Searches: Aspose.Cells C# add slicer to existing pivot table example | how to programmatically create slicer for pivot table in .NET | place Excel slicer at specific cell using Aspose.Cells | set slicer caption based on base field name Aspose.Cells | add multiple slicers for each pivot field Aspose.Cells C#
// Tags: add slicer to pivot table Aspose.Cells | link slicer with pivot table .NET | position slicer at cell E3 Aspose.Cells | customize slicer caption programmatically | create slicer for each base field Excel

using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Loads input.xlsx, retrieves the first pivot table, adds a slicer for its first base field at cell E3, sets a caption based on the field name, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Assume the first worksheet contains the pivot table
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve the first pivot table in the worksheet
        if (sheet.PivotTables.Count == 0)
        {
            System.Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }
        PivotTable pivot = sheet.PivotTables[0];

        // Ensure the pivot table has at least one base field to use for the slicer
        if (pivot.BaseFields.Count == 0)
        {
            System.Console.WriteLine("Pivot table has no base fields.");
            return;
        }
        string baseFieldName = pivot.BaseFields[0].Name;

        // Add a slicer linked to the pivot table at cell E3
        int slicerIndex = sheet.Slicers.Add(pivot, "E3", baseFieldName);
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Optional: set slicer properties
        slicer.Caption = $"{baseFieldName} Slicer";

        // Save the workbook with the new slicer
        workbook.Save("output.xlsx");
    }
}
