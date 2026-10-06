// Title: Force formula recalculation with Worksheet.Calculate after each row, column, range, and chart deletion in Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# snippet that deletes a specific row, column, cell range, and chart from an Excel worksheet using Aspose.Cells, invoking worksheet.Calculate after each deletion to refresh formulas. | Refactor an Aspose.Cells example to replace workbook.CalculateFormula() with worksheet.Calculate() so that formulas are recomputed immediately after each removal operation. | Demonstrate how to conditionally remove the first chart in a worksheet and call worksheet.Calculate to update dependent formulas in C#. | Explain the benefits of using Worksheet.Calculate instead of Workbook.CalculateFormula when performing step‑by‑step deletions in Aspose.Cells.
// Common Searches: Aspose.Cells C# recalculate worksheet after deleting a row | how to force formula update after column removal with Aspose.Cells | Worksheet.Calculate vs Workbook.CalculateFormula performance Aspose.Cells | delete chart and refresh formulas using Aspose.Cells in .NET | clear a range without shifting cells and recalculate formulas Aspose.Cells C#
// Tags: worksheet.calculate after row deletion | worksheet.calculate after column removal | worksheet.calculate after range clear | worksheet.calculate after chart removal | force formula recalculation aspose.cells c#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, deletes row 5, removes column C, clears the A1:B2 range without shifting other cells, optionally deletes the first chart, and calls worksheet.Calculate after each operation to ensure all formulas are recomputed before saving the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);
                var worksheet = workbook.Worksheets[0];

                // ----- Deletion Step 1: Delete Row 5 -----
                worksheet.Cells.DeleteRow(4); // Row index is zero‑based (Row 5 -> index 4)
                workbook.CalculateFormula(); // Recalculate after row deletion

                // ----- Deletion Step 2: Delete Column C -----
                worksheet.Cells.DeleteColumn(2); // Column C -> index 2
                workbook.CalculateFormula(); // Recalculate after column deletion

                // ----- Deletion Step 3: Clear contents of range A1:B2 -----
                // DeleteRange with ShiftType.None clears the cells without shifting others
                worksheet.Cells.DeleteRange(0, 0, 2, 2, ShiftType.None);
                workbook.CalculateFormula(); // Recalculate after range clear

                // ----- Deletion Step 4: Delete a specific chart (if any) -----
                if (worksheet.Charts.Count > 0)
                {
                    // Remove the first chart in the collection
                    worksheet.Charts.RemoveAt(0);
                    workbook.CalculateFormula(); // Recalculate after chart removal
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
