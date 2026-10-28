// Title: Create a shared formula for cells C3:C12 using Cell.SetSharedFormula in Aspose.Cells for .NET (C#)
// Common Searches: asp.net aspose.cells set shared formula for column C range C3:C12 | c# example using Cell.SetSharedFormula to copy formula across rows | how to define a shared formula owner in Aspose.Cells .NET | apply the same Excel formula to multiple cells with Aspose.Cells API | shared formula implementation in Aspose.Cells workbook using C#
// Tags: shared-formula SetSharedFormula Aspose.Cells | apply-shared-formula to range C3:C12 C# | Aspose.Cells shared-formula owner flag | Excel workbook shared-formula .NET | Cell.SetSharedFormula usage example

using System;
using Aspose.Cells;

// The example creates a new Workbook, defines the range C3:C12, sets a shared formula "=A3+B3" on the first cell (C3) with Cell.SetSharedFormula (owner flag = 0), then applies the same shared formula to the remaining cells in the column, and saves the file as SharedFormula.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Define the range C3:C12 (zero‑based indices)
            int startRow = 2;   // Row 3
            int endRow = 11;    // Row 12
            int column = 2;     // Column C

            // Set a shared formula on the first cell (C3)
            // The third parameter (0) indicates a shared formula owner in current API version
            var firstCell = worksheet.Cells[startRow, column];
            firstCell.SetSharedFormula("=A3+B3", 0, 0);

            // Apply the same shared formula to the remaining cells in the range
            for (int row = startRow + 1; row <= endRow; row++)
            {
                var cell = worksheet.Cells[row, column];
                // The third parameter (0) indicates this cell uses the shared formula defined earlier
                cell.SetSharedFormula(null, 0, 0);
            }

            // Save the workbook to a file
            workbook.Save("SharedFormula.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
