using System;
using Aspose.Cells;

namespace AsposeCellsFormulaCopy
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];

            // ---------- Populate source range with sample data and formulas ----------
            // Source range: A1:C3
            // A1 = 10, B1 = 20, C1 = 30
            // A2 = =A1*2   (relative reference)
            // B2 = =B1+5   (relative reference)
            // C2 = =C1-3   (relative reference)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);
            sheet.Cells["C1"].PutValue(30);
            sheet.Cells["A2"].Formula = "=A1*2";
            sheet.Cells["B2"].Formula = "=B1+5";
            sheet.Cells["C2"].Formula = "=C1-3";

            // Destination range starts at E1 (same size as source)
            int destFirstRow = 0;   // row index for E1
            int destFirstColumn = 4; // column index for E (0‑based)

            // ---------- Copy formulas with conversion to absolute references ----------
            // Iterate through each cell in the source range
            for (int row = 0; row < 2; row++)          // only rows with formulas (A2:C2)
            {
                for (int col = 0; col < 3; col++)      // columns A‑C
                {
                    Cell srcCell = sheet.Cells[row, col];
                    if (!string.IsNullOrEmpty(srcCell.Formula))
                    {
                        // Convert the formula to absolute A1 style.
                        // First convert to R1C1 using the source cell as base,
                        // then convert back to A1 using the destination cell as base.
                        // This yields absolute references relative to the destination.
                        string r1c1 = sheet.ConvertFormulaReferenceStyle(srcCell.Formula, true,
                                                                         srcCell.Row, srcCell.Column);
                        string absoluteA1 = sheet.ConvertFormulaReferenceStyle(r1c1, false,
                                                                                destFirstRow + row,
                                                                                destFirstColumn + col);

                        // Set the absolute formula into the destination cell
                        Cell destCell = sheet.Cells[destFirstRow + row, destFirstColumn + col];
                        destCell.SetFormula(absoluteA1, new FormulaParseOptions());

                        // Optional: copy the calculated value immediately
                        // (useful if you want values without recalculating later)
                        destCell.Value = srcCell.Value;
                    }
                }
            }

            // ---------- Save the workbook ----------
            wb.Save("FormulaCopyAbsolute.xlsx");
        }
    }
}