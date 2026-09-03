// Title: How to remove formatting‑only cells from every worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loops through all worksheets and deletes cells that contain only formatting, leaving the workbook smaller. | Create a reusable method that accepts input and output file paths, finds cells where CellValueType.IsNull, and empties them with PutValue(null). | Demonstrate how to gather formatting‑only cells into a collection and clear them efficiently to reduce file size in Aspose.Cells.
// Common Searches: asp.net remove cells that only have formatting from Excel workbook | Aspose.Cells delete empty styled cells to shrink file size | C# clear cells with CellValueType.IsNull across all worksheets | optimize Excel file size by cleaning up formatting‑only cells using Aspose.Cells | batch remove formatting‑only cells from large workbook .NET
// Tags: strip formatting‑only cells Aspose.Cells | clear IsNull cells across worksheets | reduce Excel workbook size with cell cleanup | Aspose.Cells delete empty styled cells | batch clear null‑type cells .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates each worksheet, identifies cells whose Type is IsNull (indicating no data, only formatting), sets their value to null to clear them, and saves the cleaned workbook, resulting in a smaller file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Collect cells that contain no data (only formatting)
                List<Cell> cellsToClear = new List<Cell>();
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.Type == CellValueType.IsNull)
                    {
                        cellsToClear.Add(cell);
                    }
                }

                // Clear the collected cells (remove any residual data)
                foreach (Cell cell in cellsToClear)
                {
                    // Setting the value to null ensures the cell is truly empty
                    cell.PutValue(null);
                }
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
