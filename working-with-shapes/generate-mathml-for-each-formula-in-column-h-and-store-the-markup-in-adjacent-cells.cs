using System;
using Aspose.Cells;

namespace AsposeCellsMathMLExport
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the range of used rows in column H (zero‑based index 7)
            int maxRow = cells.MaxDataRow;
            for (int row = 0; row <= maxRow; row++)
            {
                // Get the cell in column H
                Cell formulaCell = cells[row, 7]; // Column H

                // Process only if the cell contains a formula
                if (!string.IsNullOrEmpty(formulaCell.Formula))
                {
                    string formula = formulaCell.Formula;

                    // Simple conversion: wrap the formula string in MathML <math> tags.
                    // For real scenarios, replace this with a proper conversion routine.
                    string mathML = $"<math>{System.Security.SecurityElement.Escape(formula)}</math>";

                    // Write the MathML string to the adjacent cell in column I (index 8)
                    cells[row, 8].PutValue(mathML);
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output_with_mathml.xlsx");
        }
    }
}