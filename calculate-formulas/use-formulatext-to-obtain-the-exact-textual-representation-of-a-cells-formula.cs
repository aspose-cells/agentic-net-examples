// Title: C# – Retrieve Exact Formula Text from a Cell with Aspose.Cells GetFormula (non‑localized)
// Description: Demonstrates how to create a workbook, insert values, assign a formula to a cell, and use `GetFormula(false, false)` to obtain the raw A1‑notation string of the formula. The example also shows how to display formulas with `ShowFormulas` and save the file.
// Keywords: Aspose.Cells C# GetFormula | retrieve cell formula text | non‑localized formula string | Aspose.Cells FormulaText example | show formulas worksheet | extract formula A1 notation | Aspose.Cells .NET tutorial
// Common Searches: How to get the exact formula string from a cell using Aspose.Cells for .NET | Aspose.Cells GetFormula false false non‑localized formula | Display formulas instead of values in an Aspose.Cells worksheet | C# code to read cell formula as text with Aspose.Cells | Aspose.Cells extract raw formula text
// Developer Intent: Obtain the precise, non‑localized textual representation of a cell's formula.
// Use Cases: Log or audit formulas in a workbook for compliance or debugging. | Compare formulas across multiple cells to detect inconsistencies before data processing. | Generate documentation that lists each cell's formula in a readable, language‑independent format.
// AI Prompts: Write C# code using Aspose.Cells to retrieve and log the non‑localized formula of cell D5. | Provide an example that toggles ShowFormulas for a worksheet and extracts formula texts for a range of cells. | Show how to export extracted formula strings to a CSV file for further analysis.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaTextDemo
{
    // Demonstrates how to create a workbook, insert values, assign a formula to a cell, and use `GetFormula(false, false)` to obtain the raw A1‑notation string of the formula. The example also shows how to display formulas with `ShowFormulas` and save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(20);

            // Set a formula in cell C1
            Cell formulaCell = cells["C1"];
            formulaCell.Formula = "=A1+B1";

            // Obtain the exact textual representation of the formula
            // GetFormula(false, false) returns the formula in A1 notation, non‑localized
            string formulaText = formulaCell.GetFormula(false, false);

            // Display the retrieved formula text
            Console.WriteLine("Exact formula text in C1: " + formulaText);

            // Optionally, show the formula directly by enabling ShowFormulas
            worksheet.ShowFormulas = true;
            Console.WriteLine("Cell C1 displayed as formula: " + cells["C1"].StringValue);

            // Save the workbook (optional)
            workbook.Save("FormulaTextDemo.xlsx");
        }
    }
}
