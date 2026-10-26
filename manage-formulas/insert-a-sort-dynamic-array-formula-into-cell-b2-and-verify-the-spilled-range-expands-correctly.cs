// Title: How to insert a SORT dynamic array formula in cell B2 and retrieve its spilled range using Aspose.Cells for .NET
// Common Searches: C# code to apply SORT dynamic array formula with Aspose.Cells and get spill range | Aspose.Cells .NET determine spilled area of a dynamic array formula | Retrieve address of SORT formula spill range in Excel using Aspose.Cells | Calculate dynamic array results and find non‑empty cells with Aspose.Cells C# | Example of inserting SORT(A2:A5) and reading spilled range in Aspose.Cells
// Tags: SORT dynamic array Aspose.Cells .NET | spilled range detection Aspose.Cells | dynamic array formula evaluation C# | Excel formula calculation Aspose.Cells | retrieve spill address C#

using System;
using Aspose.Cells;

// The example creates a workbook, fills A2:A5 with numbers, inserts the SORT(A2:A5) dynamic array formula into B2, calculates the workbook, then scans column B until an empty cell is found to build the spilled range address (e.g., B2:B5) and outputs it.
class SortDynamicArrayExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data in column A (A2:A5)
            sheet.Cells["A2"].PutValue(5);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["A4"].PutValue(8);
            sheet.Cells["A5"].PutValue(1);

            // Insert SORT dynamic array formula into B2 (ascending order)
            sheet.Cells["B2"].Formula = "=SORT(A2:A5)";

            // Calculate the workbook to evaluate the formula
            workbook.CalculateFormula();

            // Determine the spilled range of the dynamic array formula manually
            int startRow = 1;               // Row index for B2 (zero‑based)
            int columnIndex = 1;            // Column index for column B (zero‑based)
            int currentRow = startRow;

            // Scan downwards until an empty cell is encountered
            while (true)
            {
                Cell c = sheet.Cells[currentRow, columnIndex];
                // In Aspose.Cells, a cell is considered empty when its Value is null
                if (c.Value == null)
                    break;
                currentRow++;
            }

            // The last non‑empty row is one before the empty cell
            int endRow = currentRow - 1;

            // Build the address string (Excel uses 1‑based row numbers)
            string spilledRangeAddress = $"B{startRow + 1}:B{endRow + 1}";
            Console.WriteLine("Spilled range address: " + spilledRangeAddress);
            // Expected output: Spilled range address: B2:B5

            // (Optional) Save the workbook to inspect the result manually
            // workbook.Save("SortDynamicArray.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
