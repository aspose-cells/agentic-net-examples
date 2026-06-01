using System;
using Aspose.Cells;

namespace AsposeCellsMergeFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data that will be summed
            // Row 2 (index 1) values in columns A to D
            cells[1, 0].PutValue(10); // A2
            cells[1, 1].PutValue(20); // B2
            cells[1, 2].PutValue(30); // C2
            cells[1, 3].PutValue(40); // D2

            // Merge cells across columns A to D in the first row (A1:D1)
            // Parameters: firstRow = 0, firstColumn = 0, totalRows = 1, totalColumns = 4
            cells.Merge(0, 0, 1, 4);

            // Set a formula in the merged cell (addressed by the upper‑left cell A1)
            // The formula calculates the total of the values placed in row 2
            cells[0, 0].Formula = "=SUM(A2:D2)";

            // Recalculate formulas so the result is stored
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("MergedCellFormula.xlsx");
        }
    }
}