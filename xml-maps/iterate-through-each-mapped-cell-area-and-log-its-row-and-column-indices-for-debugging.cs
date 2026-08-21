// Title: C# – Enumerate Merged Cell Areas and Log Row/Column Indices with Aspose.Cells
// Description: Demonstrates how to create a workbook, merge ranges (e.g., A1:C3, B4:E4), retrieve all merged CellArea objects using GetMergedAreas(), iterate through them, and output each area's StartRow, StartColumn, EndRow, and EndColumn for debugging. The workbook can then be saved.
// Keywords: Aspose.Cells | .NET | C# merged cells | CellArea | GetMergedAreas | debug merged ranges | enumerate merged areas | worksheet cell coordinates | XML map validation | Aspose.Cells example
// Common Searches: Aspose.Cells get merged cell ranges .NET | How to list merged cells with start and end rows in C# | Iterate over CellArea array Aspose.Cells | Debug merged cell coordinates Aspose.Cells | Retrieve merged areas from worksheet using Aspose
// Developer Intent: Extract and display the row and column indices of every merged cell region in a worksheet to aid debugging and validation.
// Use Cases: Validate programmatically merged regions during workbook generation. | Cross‑check XML map cell references against actual merged areas. | Produce a diagnostic report of all merged ranges for quality assurance.
// AI Prompts: Generate C# code that uses Aspose.Cells to list all merged cell areas and print their start/end row and column indices. | Show how to obtain CellArea objects from a worksheet and output their coordinates for debugging purposes. | Write a method that returns merged range strings (e.g., "A1:C3") from CellArea properties in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, merge ranges (e.g., A1:C3, B4:E4), retrieve all merged CellArea objects using GetMergedAreas(), iterate through them, and output each area's StartRow, StartColumn, EndRow, and EndColumn for debugging. The workbook can then be saved.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Sample merged cell areas (for demonstration)
        cells.Merge(0, 0, 2, 2); // Merge A1:C3
        cells.Merge(3, 1, 1, 4); // Merge B4:E4

        // Retrieve all merged cell areas
        CellArea[] mergedAreas = cells.GetMergedAreas();

        // Iterate through each cell area and log its row/column indices
        for (int i = 0; i < mergedAreas.Length; i++)
        {
            CellArea area = mergedAreas[i];
            Console.WriteLine(
                $"Area {i}: StartRow={area.StartRow}, StartColumn={area.StartColumn}, " +
                $"EndRow={area.EndRow}, EndColumn={area.EndColumn}");
        }

        // Save the workbook (optional, just to demonstrate lifecycle usage)
        workbook.Save("MappedCellAreasDemo.xlsx");
    }
}
