// Title: C# – Retrieve Exact Formula Text with Aspose.Cells FORMULATEXT
// Description: Demonstrates how to set a formula in a workbook, use Aspose.Cells CalculateFormula with the Excel FORMULATEXT function, and obtain the formula string for a cell, then display and save the result.
// Keywords: Aspose.Cells | C# | FORMULATEXT | CalculateFormula | read cell formula | extract formula string | Excel API example | formula text retrieval
// Common Searches: Aspose.Cells FORMULATEXT C# example | how to get formula as text with Aspose.Cells | C# retrieve Excel formula string programmatically | using CalculateFormula to read cell formula | extract original formula from workbook
// Developer Intent: Get the literal formula text from a worksheet cell using Aspose.Cells.
// Use Cases: Show formulas in UI for audit trails | Log or export formulas before bulk edits | Validate that calculated results match the defined expressions
// AI Prompts: Generate code that extracts formula text from a range of cells in a loop. | Explain error handling when FORMULATEXT is called on a cell without a formula. | Create a comparison routine that checks if the stored formula matches the FORMULATEXT output.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaTextDemo
{
    // Demonstrates how to set a formula in a workbook, use Aspose.Cells CalculateFormula with the Excel FORMULATEXT function, and obtain the formula string for a cell, then display and save the result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set a formula in cell A1
            cells["A1"].Formula = "=SUM(B1:B3)";

            // Populate the cells referenced by the formula
            cells["B1"].PutValue(10);
            cells["B2"].PutValue(20);
            cells["B3"].PutValue(30);

            // Use the Excel FORMULATEXT function via CalculateFormula to retrieve the exact formula text
            object formulaText = sheet.CalculateFormula("=FORMULATEXT(A1)");

            // Display the original formula and the text obtained from FORMULATEXT
            Console.WriteLine("Original formula in A1: " + cells["A1"].Formula);
            Console.WriteLine("Exact textual representation via FORMULATEXT: " + formulaText);

            // Save the workbook (optional)
            workbook.Save("FormulaTextDemo.xlsx");
        }
    }
}
