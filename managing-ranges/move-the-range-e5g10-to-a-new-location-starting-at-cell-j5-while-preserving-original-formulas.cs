// Title: C# – Move Excel range E5:G10 to J5 while preserving formulas using AspNet Aspose.Cells
// Description: Demonstrates how to load a workbook, define the source CellArea (E5:G10), relocate the range to J5 with Cells.MoveRange, keep all formulas intact, and save the result as a new file.
// Keywords: Aspose.Cells MoveRange C# | move Excel range preserving formulas | relocate cell block E5:G10 to J5 | Aspose.Cells shift range | C# Excel range move example
// Common Searches: Aspose.Cells move range without losing formulas | C# MoveRange E5:G10 to J5 | how to shift a block of cells in Aspose.Cells | preserve formulas when moving Excel range C# | Aspose.Cells MoveRange method usage
// Developer Intent: Move the cells in the range E5:G10 to start at J5, ensuring that all existing formulas remain functional after the relocation.
// Use Cases: Reposition a calculated table in a financial model without breaking dependent formulas. | Shift a data block that feeds a chart to a new column while keeping references accurate. | Programmatically reorganize a template section during dynamic report generation, preserving formula logic.
// AI Prompts: Write C# code with Aspose.Cells to move range D2:F7 to H2 and keep formulas unchanged. | Explain how Cells.MoveRange treats relative and absolute references in Aspose.Cells. | Provide a guide for moving multiple non‑contiguous ranges in a worksheet while retaining their formulas.

using System;
using Aspose.Cells;

// Demonstrates how to load a workbook, define the source CellArea (E5:G10), relocate the range to J5 with Cells.MoveRange, keep all formulas intact, and save the result as a new file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define the source range E5:G10 (zero‑based indices)
        CellArea sourceArea = new CellArea
        {
            StartRow = 4,    // Row 5
            StartColumn = 4, // Column E
            EndRow = 9,      // Row 10
            EndColumn = 6    // Column G
        };

        // Move the range to start at J5 (row 5, column J)
        // Destination row index = 4 (row 5), destination column index = 9 (column J)
        cells.MoveRange(sourceArea, 4, 9);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
