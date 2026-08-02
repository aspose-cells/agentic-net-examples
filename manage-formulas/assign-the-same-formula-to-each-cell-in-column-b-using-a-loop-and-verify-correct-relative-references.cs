// Title: Assign a single R1C1 formula to an entire column with a loop and verify relative references using Aspose.Cells for .NET
// Description: This C# example creates a workbook, fills column A with numbers 1‑10, loops through rows to set the R1C1 formula "=RC[-1]" in each cell of column B, recalculates the sheet, prints a verification that every B cell matches its adjacent A cell, and saves the file as ColumnBFormulaLoopDemo.xlsx.
// Keywords: Aspose.Cells C# set column formula loop | R1C1Formula batch assignment | verify relative references Aspose.Cells | calculate workbook formulas .NET | save Excel file after formula update
// Common Searches: apply same R1C1 formula to a whole column Aspose.Cells | loop to assign formulas in C# Aspose.Cells example | check that column B equals column A after formula calculation | batch set formulas in Excel using Aspose.Cells for .NET
// Developer Intent: Use a for‑loop to assign the R1C1 formula "=RC[-1]" to every cell in column B, recalculate the workbook, and confirm that each B cell correctly references its left‑hand neighbor in column A.
// Use Cases: Copy values from one column to another with a relative reference formula. | Populate large spreadsheets with identical calculations without manual entry. | Validate formula outcomes against source data before exporting the workbook.
// AI Prompts: Write C# code with Aspose.Cells that sets the formula '=RC[-1]' in column B for rows 1‑100 and verifies each result against column A. | Explain how to loop through rows in Aspose.Cells, assign a relative reference formula, recalculate the sheet, and report any mismatches. | Provide a step‑by‑step tutorial for assigning a single formula to a column, executing calculations, and saving the workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLoopDemo
{
    // This C# example creates a workbook, fills column A with numbers 1‑10, loops through rows to set the R1C1 formula "=RC[-1]" in each cell of column B, recalculates the sheet, prints a verification that every B cell matches its adjacent A cell, and saves the file as ColumnBFormulaLoopDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A with sample data (1 to 10)
            for (int row = 0; row < 10; row++)
            {
                cells[row, 0].PutValue(row + 1); // A1..A10
            }

            // Assign the same R1C1 formula to each cell in column B using a loop.
            // The formula "=RC[-1]" means "current row, one column to the left",
            // so each B cell will reference its corresponding A cell.
            for (int row = 0; row < 10; row++)
            {
                cells[row, 1].R1C1Formula = "=RC[-1]"; // B column
            }

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Verify that each B cell value equals the value in column A
            Console.WriteLine("Verification of relative references (A vs B):");
            for (int row = 0; row < 10; row++)
            {
                object aVal = cells[row, 0].Value;
                object bVal = cells[row, 1].Value;
                Console.WriteLine($"Row {row + 1}: A = {aVal}, B = {bVal} (Match: {aVal.Equals(bVal)})");
            }

            // Save the workbook
            workbook.Save("ColumnBFormulaLoopDemo.xlsx");
        }
    }
}
