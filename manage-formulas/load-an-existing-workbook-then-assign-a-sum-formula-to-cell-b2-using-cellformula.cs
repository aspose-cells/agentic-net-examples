// Title: Set a SUM Formula in Cell B2 of an Existing Workbook with Aspose.Cells for .NET (C#)
// Description: Loads an existing Excel file, accesses the first worksheet, assigns the formula "=SUM(A1:A5)" to cell B2, forces calculation of all formulas, and saves the updated workbook as a new file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# set formula | load workbook Aspose.Cells | assign SUM formula Excel | calculate formulas Aspose | save workbook C#
// Common Searches: Aspose.Cells set cell formula C# | How to add SUM formula to Excel file using Aspose.Cells | Recalculate workbook after adding formula Aspose.Cells | Save modified Excel with Aspose.Cells .NET
// Developer Intent: Load a workbook, place a SUM formula in B2, evaluate it, and write the changes back to disk.
// Use Cases: Generate a totals row in a report template automatically. | Update a shared spreadsheet with dynamic calculations before distribution. | Insert calculation formulas during bulk data import for financial models.
// AI Prompts: Write C# code that uses Aspose.Cells to set a custom formula in cell C3 based on a user‑defined range. | Show how to copy the same SUM formula down a column and recalculate the workbook with Aspose.Cells. | Explain handling of empty or non‑numeric cells in the source range when applying a SUM formula with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    // Loads an existing Excel file, accesses the first worksheet, assigns the formula "=SUM(A1:A5)" to cell B2, forces calculation of all formulas, and saves the updated workbook as a new file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing workbook
            string inputPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the target cell B2
            Cell targetCell = worksheet.Cells["B2"];

            // Assign a SUM formula to B2 (e.g., sum of A1 through A5)
            targetCell.Formula = "=SUM(A1:A5)";

            // Optional: calculate the formula so the result is stored in the cell
            workbook.CalculateFormula();

            // Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Formula assigned to B2 and workbook saved to '{outputPath}'.");
        }
    }
}
