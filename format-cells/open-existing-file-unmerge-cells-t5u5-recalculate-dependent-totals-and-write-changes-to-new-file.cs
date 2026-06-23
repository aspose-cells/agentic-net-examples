using System;
using Aspose.Cells;

namespace AsposeCellsUnmergeAndRecalculate
{
    class Program
    {
        static void Main()
        {
            // Paths to the source and destination Excel files
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Unmerge the range T5:U5
            // Column T = 19 (zero‑based), U = 20, Row 5 = index 4
            // Unmerge 1 row and 2 columns
            cells.UnMerge(firstRow: 4, firstColumn: 19, totalRows: 1, totalColumns: 2);

            // Recalculate all formulas so dependent totals are updated
            workbook.CalculateFormula();

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
        }
    }
}