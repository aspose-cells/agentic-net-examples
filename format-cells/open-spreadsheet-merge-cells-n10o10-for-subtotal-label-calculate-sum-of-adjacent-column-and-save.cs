// Title: C# – Merge N10:O10, add Subtotal label with SUM formula, recalculate and save using Aspose.Cells
// Description: Loads an Excel file, merges cells N10:O10 on the first worksheet, writes "Subtotal", inserts the formula =SUM(N1:N9), forces formula evaluation, and saves the updated workbook.
// Keywords: Aspose.Cells C# merge cells | subtotal label Excel | SUM formula Aspose.Cells | recalculate formulas .NET | save workbook after merge | Excel automation C#
// Common Searches: Aspose.Cells merge N10 O10 C# | how to add subtotal label with SUM formula using Aspose.Cells | recalculate formulas after merging cells Aspose.Cells .NET | save workbook after inserting formula Aspose.Cells | C# example for merging cells and summing a column in Excel
// Developer Intent: Merge a specific cell range, place a subtotal label with a SUM formula, evaluate the formula, and save the workbook.
// Use Cases: Create a subtotal row in a financial report by merging N10:O10, labeling it, and summing N1:N9. | Automate column totals before exporting data to another system. | Generate a consolidated total in an invoice template with a merged label cell and dynamic sum.
// AI Prompts: Provide C# code that merges cells N10:O10, writes "Subtotal", adds =SUM(N1:N9), recalculates formulas, and saves the file using Aspose.Cells. | Show how to force formula calculation after merging cells and inserting a subtotal in Aspose.Cells for .NET. | Explain how to style the merged subtotal cell (e.g., bold, background color) after adding a SUM formula with Aspose.Cells.

using Aspose.Cells;
using System;

// Loads an Excel file, merges cells N10:O10 on the first worksheet, writes "Subtotal", inserts the formula =SUM(N1:N9), forces formula evaluation, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Paths for input and output files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputPath);
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells N10:O10 (row 9, column 13, 1 row, 2 columns)
        cells.Merge(9, 13, 1, 2);

        // Set a label in the merged cell (optional)
        cells["N10"].PutValue("Subtotal");

        // Insert a formula that sums the values in column N from rows 1 to 9
        cells["N10"].Formula = "=SUM(N1:N9)";

        // Recalculate formulas so the sum is evaluated
        workbook.CalculateFormula();

        // Save the modified workbook
        workbook.Save(outputPath);
    }
}
