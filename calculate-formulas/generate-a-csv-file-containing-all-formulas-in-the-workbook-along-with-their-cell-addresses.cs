using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaExport
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (replace with your actual file path)
            string workbookPath = "input.xlsx";
            Workbook workbook = new Workbook(workbookPath);

            // Prepare the CSV output file
            string csvPath = "formulas.csv";
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // Write CSV header
                writer.WriteLine("Worksheet,CellAddress,Formula");

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all cells that contain data
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            // Check if the cell has a formula
                            if (!string.IsNullOrEmpty(cell.Formula))
                            {
                                // Get the cell address in A1 style
                                string address = cell.Name;
                                // Escape double quotes in the formula for CSV compliance
                                string formula = cell.Formula.Replace("\"", "\"\"");
                                // Write a CSV line: Worksheet name, cell address, formula
                                writer.WriteLine($"\"{sheet.Name}\",\"{address}\",\"{formula}\"");
                            }
                        }
                    }
                }
            }

            Console.WriteLine($"Formulas have been exported to '{csvPath}'.");
        }
    }
}