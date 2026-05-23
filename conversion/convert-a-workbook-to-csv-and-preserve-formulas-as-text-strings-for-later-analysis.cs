using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsCsvFormulaExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Path for the resulting CSV file
            string csvPath = "output.csv";

            // Load the workbook (lifecycle: create/load)
            Workbook originalWorkbook = new Workbook(sourcePath);

            // Create a new workbook and copy the content of the original workbook
            // (lifecycle: copy using provided rule)
            Workbook workbookForCsv = new Workbook();
            workbookForCsv.Copy(originalWorkbook);

            // Iterate through all worksheets and cells
            foreach (Worksheet sheet in workbookForCsv.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Get the maximum used row and column indices
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        // If the cell contains a formula, replace its value with the formula text
                        if (cell.IsFormula)
                        {
                            // Preserve the formula as a plain string
                            cell.PutValue(cell.Formula);
                        }
                    }
                }
            }

            // Save the modified workbook as CSV (lifecycle: save using provided rule)
            workbookForCsv.Save(csvPath, SaveFormat.Csv);

            Console.WriteLine($"Workbook converted to CSV with formulas preserved as text at: {csvPath}");
        }
    }
}