using System;
using Aspose.Cells;

namespace ReplaceTodayWithStaticDate
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Path where the modified workbook will be saved
            string outputPath = "output.xlsx";

            // Snapshot date that will replace all TODAY() occurrences
            DateTime snapshotDate = new DateTime(2023, 5, 1);

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to limit iteration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan all cells within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Identify cells that contain a formula with TODAY()
                        if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula) &&
                            cell.Formula.IndexOf("TODAY()", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Replace the formula with the static snapshot date
                            // PutValue overwrites the formula and stores the date as a value
                            cell.PutValue(snapshotDate);
                        }
                    }
                }
            }

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputPath);
        }
    }
}