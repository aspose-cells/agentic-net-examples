using System;
using Aspose.Cells;

class CheckMergedCells
{
    static void Main()
    {
        // Load an existing workbook (replace "input.xlsx" with your file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve all merged areas in the worksheet
        CellArea[] mergedAreas = worksheet.Cells.GetMergedAreas();

        // Determine whether any merged cells exist
        bool hasMerged = mergedAreas != null && mergedAreas.Length > 0;
        Console.WriteLine("Worksheet has merged cells: " + hasMerged);

        // Optional: display details of each merged area
        foreach (CellArea area in mergedAreas)
        {
            Console.WriteLine($"Merged area: Row[{area.StartRow}-{area.EndRow}], Column[{area.StartColumn}-{area.EndColumn}]");
        }

        // Save the workbook if further processing is required
        workbook.Save("output.xlsx");
    }
}