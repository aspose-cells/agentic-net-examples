using System;
using Aspose.Cells;

namespace AsposeCellsOffsetFormulaScanner
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be scanned
            string inputPath = "input.xlsx";

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to avoid scanning empty cells
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only formula cells
                        if (cell.IsFormula)
                        {
                            // Check if the formula uses the OFFSET function (case‑insensitive)
                            if (cell.Formula.IndexOf("OFFSET", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // OFFSET is a volatile function, so we note its volatility
                                Console.WriteLine($"Worksheet: {sheet.Name}, Cell: {cell.Name}, Formula: {cell.Formula} (volatile)");
                            }
                        }
                    }
                }
            }

            // Optionally, save the workbook after processing (save rule)
            // workbook.Save("output.xlsx");
        }
    }
}