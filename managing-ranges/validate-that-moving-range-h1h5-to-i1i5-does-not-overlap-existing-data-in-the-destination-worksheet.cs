// Title: Validate that the destination range I1:I5 is empty before moving range H1:H5 with Aspose.Cells for .NET
// Common Searches: how to ensure I1:I5 is empty before moving H1:H5 using Aspose.Cells C# | Aspose.Cells .NET check for overlapping cells when shifting a column range | prevent data overwrite when moving a column range in Excel with Aspose.Cells | C# code to verify destination range emptiness before copying cells in Aspose.Cells | move cells without overlap Aspose.Cells workbook example
// Tags: validate destination range emptiness Aspose.Cells | move column range without overlap C# | check overlapping cells Aspose.Cells .NET | copy cell style while moving range Aspose.Cells | clear source range after moving cells Aspose.Cells | column letter to index conversion CellsHelper

using Aspose.Cells;
using System;
using System.IO;

// The example loads or creates a workbook, converts column letters H and I to zero‑based indexes, checks cells I1‑I5 for existing values, and if the range is empty copies each cell's value and style from H1‑H5 to I1‑I5, clears the original H column range, and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Load the workbook if the input file exists; otherwise create a new workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. A new workbook will be created.");
                workbook = new Workbook();
                workbook.Worksheets.Add();
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Convert column letters to zero‑based indexes.
            int srcColumn = CellsHelper.ColumnNameToIndex("H"); // column H
            int destColumn = CellsHelper.ColumnNameToIndex("I"); // column I
            int startRow = 0; // row 1 (zero‑based)
            int endRow = 4;   // row 5 (zero‑based)

            // Verify that the destination range is empty.
            bool overlap = false;
            for (int row = startRow; row <= endRow; row++)
            {
                if (sheet.Cells[row, destColumn].Value != null)
                {
                    overlap = true;
                    break;
                }
            }

            if (overlap)
            {
                Console.WriteLine("Destination range I1:I5 already contains data. Move aborted.");
            }
            else
            {
                // Move each cell from source to destination.
                for (int row = startRow; row <= endRow; row++)
                {
                    Cell srcCell = sheet.Cells[row, srcColumn];
                    Cell destCell = sheet.Cells[row, destColumn];

                    // Copy value.
                    destCell.PutValue(srcCell.Value);

                    // Copy style (optional).
                    destCell.SetStyle(srcCell.GetStyle());
                }

                // Clear the original source range.
                sheet.Cells.ClearRange(startRow, srcColumn, endRow - startRow + 1, 1);

                Console.WriteLine("Range H1:H5 moved to I1:I5 successfully.");
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
