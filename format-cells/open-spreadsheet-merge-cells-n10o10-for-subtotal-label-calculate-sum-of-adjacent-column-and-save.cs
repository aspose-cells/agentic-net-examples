// Title: Merge cells N10:O10, add a 'Subtotal' label with SUM(N1:N9) formula, and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Merge the range N10:O10 on the first worksheet, set its value to "Subtotal", assign the formula SUM(N1:N9), evaluate all formulas, and save the workbook to a new file with Aspose.Cells in C#. | Load an existing Excel file, combine cells N10 and O10, write a subtotal header, embed a SUM formula that references N1:N9, force formula calculation, and persist the changes using Aspose.Cells for .NET. | Using Aspose.Cells for C#, create a merged cell for a subtotal label, apply a column‑total formula, calculate the result immediately, and export the updated workbook.
// Common Searches: Aspose.Cells C# merge N10 O10 and insert SUM formula for column total | how to set a subtotal label in a merged cell with Aspose.Cells .NET | force formula evaluation before saving workbook using Aspose.Cells C#
// Tags: merge cells Aspose.Cells C# | subtotal label formula Aspose.Cells | column sum calculation Aspose.Cells | force formula evaluation Aspose.Cells | save modified workbook Aspose.Cells C#

using System;
using Aspose.Cells;

// The example loads a workbook, merges cells N10:O10 on the first worksheet, writes "Subtotal" in the merged range, assigns a SUM(N1:N9) formula, forces formula calculation, and saves the updated file to a new location using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = @"C:\Temp\SourceWorkbook.xlsx";
        string outputPath = @"C:\Temp\ResultWorkbook.xlsx";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells N10:O10 for the subtotal label
        // N = column 14, O = column 15 (zero‑based index: 13 and 14)
        sheet.Cells.Merge(9, 13, 1, 2); // Row 9 (10th row), Column 13 (N), 1 row, 2 columns

        // Set the label text in the merged cell
        sheet.Cells["N10"].PutValue("Subtotal");

        // Calculate the sum of the adjacent column (column N, rows 1‑9) and place the result in the merged cell
        // The formula will automatically compute the sum when the workbook is opened
        sheet.Cells["N10"].Formula = "SUM(N1:N9)";

        // Optionally, you can calculate the formula now so the value is stored in the file
        workbook.CalculateFormula();

        // Save the modified workbook
        workbook.Save(outputPath);
    }
}
