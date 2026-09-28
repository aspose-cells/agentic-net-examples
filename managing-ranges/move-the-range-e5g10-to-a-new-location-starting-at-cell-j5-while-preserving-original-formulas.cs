// Title: Move Excel range E5:G10 to start at J5 while preserving formulas with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that relocates the range E5:G10 to begin at J5 and keeps all original formulas intact. | Show how to apply the MoveRange method to shift a block of cells from E5:G10 to J5 without breaking formula references in a .NET workbook.
// Common Searches: Aspose.Cells C# move range preserve formulas | How to shift Excel cells E5:G10 to J5 using Aspose.Cells .NET | Keep formula references when moving a cell block with Aspose.Cells MoveRange | C# example for relocating a range in an existing workbook with Aspose.Cells
// Tags: Aspose.Cells MoveRange method | move Excel range preserving formulas .NET | relocate cell block E5:G10 to J5 Aspose | C# shift range without breaking references

using Aspose.Cells;
using System;
using System.IO;

// The program loads "input.xlsx", uses Aspose.Cells' MoveRange to relocate the cell block E5:G10 to start at J5 on the first worksheet while retaining all formula references, and saves the result as "output.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the source range to move (E5:G10)
            CellArea sourceRange = CellArea.CreateCellArea("E5", "G10");

            // Move the range to start at J5 (row index 4, column index 9)
            sheet.Cells.MoveRange(sourceRange, 4, 9);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
