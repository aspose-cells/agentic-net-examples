// Title: Get Spilled Range of a Dynamic Array Formula in Cell C3 with Aspose.Cells for .NET
// Description: Loads an Excel workbook, accesses cell C3, verifies the presence of a dynamic array formula using IsDynamicArrayFormula, obtains the spill area via GetArrayRange, converts the resulting CellArea to a standard address (e.g., "C3:E5"), and outputs the range.
// Keywords: Aspose.Cells | GetArrayRange | dynamic array | spilled range | C# Excel | CellArea | IsDynamicArrayFormula | .NET
// Common Searches: Aspose.Cells get spilled range | C# retrieve dynamic array spill area | GetArrayRange example Aspose | find spilled range of Excel dynamic array using Aspose | Cell.IsDynamicArrayFormula usage
// Developer Intent: Determine the address of the range that a dynamic array formula in cell C3 expands to.
// Use Cases: Confirm that a dynamic array formula produced the expected spill area before further processing. | Log or display the spill address for debugging or reporting. | Apply formatting, validation, or additional calculations to the entire spilled range.
// AI Prompts: Write C# code with Aspose.Cells to check if cell C3 contains a dynamic array formula and return its spilled range address. | Create a reusable method that accepts any cell reference and returns the spilled range string for a dynamic array formula. | Explain how GetArrayRange and CellArea work together to retrieve spill addresses in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Loads an Excel workbook, accesses cell C3, verifies the presence of a dynamic array formula using IsDynamicArrayFormula, obtains the spill area via GetArrayRange, converts the resulting CellArea to a standard address (e.g., "C3:E5"), and outputs the range.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or specify the appropriate one)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the cell C3 which contains the dynamic array formula
        Cell dynamicCell = worksheet.Cells["C3"];

        // Verify that the cell indeed holds a dynamic array formula
        if (dynamicCell.IsDynamicArrayFormula)
        {
            // Retrieve the spilled range of the dynamic array formula
            CellArea spillArea = dynamicCell.GetArrayRange();

            // Convert the CellArea to a readable address (e.g., "C3:E5")
            string startAddress = worksheet.Cells[spillArea.StartRow, spillArea.StartColumn].Name;
            string endAddress   = worksheet.Cells[spillArea.EndRow,   spillArea.EndColumn].Name;
            string spilledRange = $"{startAddress}:{endAddress}";

            Console.WriteLine($"Spilled range for C3: {spilledRange}");
        }
        else
        {
            Console.WriteLine("Cell C3 does not contain a dynamic array formula.");
        }
    }
}
