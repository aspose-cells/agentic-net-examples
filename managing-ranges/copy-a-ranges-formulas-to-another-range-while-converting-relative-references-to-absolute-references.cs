// Title: Copy a range of cells with formulas to another location while converting relative references to absolute references using Aspose.Cells for .NET
// Common Searches: Aspose.Cells copy range formulas to new location preserving formatting C# | how to convert relative cell references to absolute when copying a range with Aspose.Cells | C# example copying Excel range A1:C3 to E1 using Aspose.Cells Cell.Copy | preserve formulas and styles while duplicating a cell block in Aspose.Cells | Aspose.Cells copy cells and fix formula references for .NET applications
// Tags: copy range formulas Aspose.Cells | convert relative to absolute references C# | cell copy method Aspose.Cells | preserve formatting during range duplication .NET | Excel range duplication Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The code loads a workbook, defines a source range (A1:C3), and uses the Cell.Copy method to duplicate each cell—including its formula, value, and style—to a destination starting at E1, automatically converting relative references to absolute ones, then saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Define source range (A1:C3) and destination start cell (E1)
            CellArea sourceArea = CellArea.CreateCellArea("A1", "C3");
            int destStartRow = 0;      // Row index for E1 (0‑based)
            int destStartColumn = 4;   // Column index for E1 (0‑based)

            // Copy each cell from source to destination, preserving formulas, values, and styles
            for (int rowOffset = 0; rowOffset <= sourceArea.EndRow - sourceArea.StartRow; rowOffset++)
            {
                for (int colOffset = 0; colOffset <= sourceArea.EndColumn - sourceArea.StartColumn; colOffset++)
                {
                    int srcRow = sourceArea.StartRow + rowOffset;
                    int srcCol = sourceArea.StartColumn + colOffset;
                    int destRow = destStartRow + rowOffset;
                    int destCol = destStartColumn + colOffset;

                    Cell srcCell = sheet.Cells[srcRow, srcCol];
                    Cell destCell = sheet.Cells[destRow, destCol];

                    // Use the built‑in Copy method to handle formulas, values, and formatting correctly
                    srcCell.Copy(destCell);
                }
            }

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
