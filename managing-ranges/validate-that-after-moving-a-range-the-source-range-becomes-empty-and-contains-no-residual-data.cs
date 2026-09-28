// Title: C# example to confirm that Aspose.Cells MoveRange empties the original A1:C3 range
// AI Prompts: Use Aspose.Cells in C# to move the range A1:C3 to F6 and programmatically assert that the source cells are now blank. | Write a C# unit test with Aspose.Cells that calls Cells.MoveRange and checks that the original cell area contains no values after the move. | Generate C# code that shifts a block of cells using MoveRange and verifies the source range is empty by inspecting each cell's Value property.
// Common Searches: Aspose.Cells C# how to verify that MoveRange clears the original cells | check if source range is empty after moving range with Aspose.Cells .NET | C# Aspose.Cells MoveRange source range residual data test
// Tags: Aspose.Cells MoveRange source clearing | C# verify empty range after MoveRange | Aspose.Cells cell area emptiness validation | MoveRange operation source cells blank .NET | Aspose.Cells range relocation test

using Aspose.Cells;
using System;

// The program creates a workbook, fills cells A1:C3 with sequential numbers, moves that range to start at F6 using Cells.MoveRange, then iterates over the original cells to ensure they contain no values, prints the result, and saves the workbook as MovedRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source range A1:C3 with sample data
            int value = 1;
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].PutValue(value++);
                }
            }

            // Define source range address
            string sourceRange = "A1:C3";

            // Convert the address string to a CellArea object required by MoveRange
            string[] rangeParts = sourceRange.Split(':');
            CellArea sourceArea = CellArea.CreateCellArea(rangeParts[0], rangeParts[1]);

            // Move the range to a new location starting at F6 (row index 5, column index 5)
            sheet.Cells.MoveRange(sourceArea, 5, 5);

            // Validate that source range cells are now empty
            bool isEmpty = true;
            for (int row = 0; row < 3 && isEmpty; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.Value != null && !string.IsNullOrEmpty(cell.StringValue))
                    {
                        isEmpty = false;
                        break;
                    }
                }
            }

            Console.WriteLine(isEmpty ? "Source range is empty after move." : "Source range still contains data.");

            // Save the workbook (optional)
            workbook.Save("MovedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
