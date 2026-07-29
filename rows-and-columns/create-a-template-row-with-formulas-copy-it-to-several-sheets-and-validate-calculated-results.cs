// Title: Copy Template Rows with Shared Formulas to Multiple Worksheets and Validate Results (Aspose.Cells for .NET)
// Description: This C# sample creates a workbook, fills column A with values 1‑5, sets a shared formula (=A1*2) in column B, copies the five‑row template to additional worksheets, runs Workbook.CalculateFormula, and prints both source and calculated values before saving the file.
// Keywords: Aspose.Cells | C# | .NET | copy rows | shared formula | Workbook.CalculateFormula | template worksheet | multiple sheets | formula propagation | Excel automation
// Common Searches: Aspose.Cells copy rows with formulas | How to duplicate a template row across worksheets in C# | Validate formula calculation after copying rows Aspose.Cells | Set shared formula in Aspose.Cells .NET | Copy rows between worksheets using Aspose.Cells
// Developer Intent: Duplicate a set of rows that contain a shared formula to several worksheets and ensure the formulas recalculate correctly.
// Use Cases: Create a master calculation sheet and replicate its rows to monthly tabs while preserving the shared formula. | Build a pricing template that multiplies quantity by a factor, copy it to region‑specific sheets, and automatically compute totals. | Generate new worksheets from a template, copy rows with embedded formulas, and run CalculateFormula to verify results.
// AI Prompts: Generate C# code with Aspose.Cells that defines a shared formula (=A1*2) in column B, copies the first five rows to three additional worksheets, recalculates all formulas, and outputs the values for verification. | Provide a method that copies rows from a template worksheet to a destination worksheet while preserving shared formulas and returns a validation report of column B values. | Explain how Workbook.CalculateFormula works after copying rows with shared formulas to ensure correct recalculation across all worksheets.

using System;
using Aspose.Cells;

namespace AsposeCellsTemplateRowExample
{
    // This C# sample creates a workbook, fills column A with values 1‑5, sets a shared formula (=A1*2) in column B, copies the five‑row template to additional worksheets, runs Workbook.CalculateFormula, and prints both source and calculated values before saving the file.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();

            // ---------- Prepare the template worksheet ----------
            Worksheet templateSheet = workbook.Worksheets[0];
            Cells tmplCells = templateSheet.Cells;

            // Fill sample data in column A (rows 1 to 5)
            for (int i = 0; i < 5; i++)
            {
                tmplCells[i, 0].PutValue(i + 1); // A1..A5 = 1..5
            }

            // Set a shared formula in column B starting from B1.
            // The formula multiplies the value in column A by 2.
            // This will propagate the formula to B1:B5.
            tmplCells[0, 1].SetSharedFormula("=A1*2", 5, 1);

            // ---------- Add additional worksheets ----------
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // ---------- Copy the template rows (including data and formulas) ----------
            // Copy rows 0..4 (5 rows) from the template sheet to each target sheet.
            CopyTemplateRows(tmplCells, sheet2.Cells);
            CopyTemplateRows(tmplCells, sheet3.Cells);

            // ---------- Calculate all formulas ----------
            workbook.CalculateFormula();

            // ---------- Validate and display results ----------
            Console.WriteLine("Validation of calculated results:");
            for (int s = 0; s < workbook.Worksheets.Count; s++)
            {
                Worksheet ws = workbook.Worksheets[s];
                Console.WriteLine($"--- {ws.Name} ---");
                for (int row = 0; row < 5; row++)
                {
                    // Column A holds the original value, Column B holds the calculated result.
                    Console.WriteLine($"Row {row + 1}: A={ws.Cells[row, 0].Value}, B={ws.Cells[row, 1].Value}");
                }
            }

            // ---------- Save the workbook ----------
            workbook.Save("TemplateRowCopyResult.xlsx");
        }

        /// <param name="source">Source Cells object (template sheet).</param>
        /// <param name="dest">Destination Cells object (target sheet).</param>
        private static void CopyTemplateRows(Cells source, Cells dest)
        {
            // Parameters: sourceCells, sourceRowIndex, destinationRowIndex, rowNumber
            // Copy rows 0‑4 (5 rows) from source to destination starting at row 0.
            dest.CopyRows(source, 0, 0, 5);
        }
    }
}
