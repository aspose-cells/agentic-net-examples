using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Example merged areas (replace with your own mapped areas as needed)
        cells.Merge(0, 0, 2, 2); // Merge A1:C3
        cells.Merge(4, 1, 1, 3); // Merge B5:E5

        // Retrieve all merged cell areas using the provided GetMergedAreas method
        CellArea[] mergedAreas = worksheet.Cells.GetMergedAreas();

        // Iterate through each CellArea and log its row/column indices for debugging
        foreach (CellArea area in mergedAreas)
        {
            Console.WriteLine(
                $"Merged Area - StartRow: {area.StartRow}, StartColumn: {area.StartColumn}, " +
                $"EndRow: {area.EndRow}, EndColumn: {area.EndColumn}");
        }

        // Save the workbook (optional, just to demonstrate lifecycle usage)
        workbook.Save("MappedCellAreasDebug.xlsx");
    }
}