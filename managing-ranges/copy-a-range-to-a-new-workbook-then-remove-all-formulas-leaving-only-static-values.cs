// Title: Copy a defined cell range to a new workbook and convert all formulas to static values with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads source.xlsx, copies the A1:C10 area to a fresh workbook while preserving styles, evaluates any formulas in the copied cells and writes the results back as constant values, then saves the file as destination.xlsx using Aspose.Cells. | Create a reusable method that accepts a source file path, destination file path, and a CellArea, copies that range into a new workbook, strips all formulas leaving only their evaluated results, and returns the saved workbook.
// Common Searches: Aspose.Cells C# copy specific range to another workbook and keep only values | How to replace formulas with their results after copying a range using Aspose.Cells .NET | Copy A1:C10 from source.xlsx to new file as static values with Aspose.Cells | Remove all formulas from a copied worksheet while preserving formatting in C# Aspose.Cells | Export Excel cell area as values only using Aspose.Cells for .NET
// Tags: copy range to new workbook Aspose.Cells | replace formulas with static values .NET | export cell area as values C# | preserve cell styles while copying Aspose.Cells | evaluate formulas to constants Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads source.xlsx, copies a defined cell area (e.g., A1:C10) into a newly created workbook, retains the original formatting, iterates over the destination cells to evaluate and replace any formulas with their calculated values, and saves the result as destination.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destinationPath = "destination.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook srcWb = new Workbook(sourcePath);
            Worksheet srcWs = srcWb.Worksheets[0];

            // Define the range to copy (example: A1:C10)
            CellArea srcArea = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 0,   // Column A (zero‑based)
                EndRow = 9,        // Row 10
                EndColumn = 2      // Column C
            };

            // Create a new workbook for the destination
            Workbook destWb = new Workbook();
            destWb.Worksheets.Clear();                     // Remove default sheet
            int destSheetIndex = destWb.Worksheets.Add(); // Add fresh sheet
            Worksheet destWs = destWb.Worksheets[destSheetIndex];

            // Manually copy the defined range from source to destination starting at A1
            int rowCount = srcArea.EndRow - srcArea.StartRow + 1;
            int colCount = srcArea.EndColumn - srcArea.StartColumn + 1;

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < colCount; j++)
                {
                    int srcRow = srcArea.StartRow + i;
                    int srcCol = srcArea.StartColumn + j;
                    int destRow = i; // start at row 0 in destination
                    int destCol = j; // start at column 0 in destination

                    Cell srcCell = srcWs.Cells[srcRow, srcCol];
                    Cell destCell = destWs.Cells[destRow, destCol];

                    // Copy value, formula and style
                    if (srcCell.IsFormula)
                    {
                        destCell.Formula = srcCell.Formula;
                    }
                    else
                    {
                        destCell.PutValue(srcCell.Value);
                    }
                    destCell.SetStyle(srcCell.GetStyle());
                }
            }

            // Replace all formulas with their evaluated static values
            int maxRow = destWs.Cells.MaxDataRow;
            int maxCol = destWs.Cells.MaxDataColumn;
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = destWs.Cells[row, col];
                    if (cell.IsFormula)
                    {
                        // Evaluate the formula and write the result back as a constant value
                        cell.PutValue(cell.Value);
                    }
                }
            }

            // Save the new workbook containing only static values
            destWb.Save(destinationPath);
            Console.WriteLine($"Destination workbook saved to: {destinationPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
