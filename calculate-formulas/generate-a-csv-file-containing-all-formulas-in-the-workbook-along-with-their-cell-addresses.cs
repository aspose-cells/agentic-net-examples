using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsFormulaExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (can be .xlsx, .xls, etc.)
            string sourcePath = "input.xlsx";

            // Path where the CSV file with formulas will be saved
            string csvOutputPath = "formulas.csv";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // StringBuilder to accumulate CSV lines
            StringBuilder csvBuilder = new StringBuilder();

            // Optional: add header line
            csvBuilder.AppendLine("CellAddress,Formula");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the cells collection of the current worksheet
                Cells cells = sheet.Cells;

                // Iterate through each cell that contains data
                foreach (Cell cell in cells)
                {
                    // Check if the cell has a formula (non‑empty string)
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Escape double quotes in the formula for CSV compliance
                        string escapedFormula = cell.Formula.Replace("\"", "\"\"");

                        // Write the cell address (e.g., A1) and its formula
                        csvBuilder.AppendLine($"{cell.Name},\"{escapedFormula}\"");
                    }
                }
            }

            // Write the accumulated CSV content to the output file
            File.WriteAllText(csvOutputPath, csvBuilder.ToString(), Encoding.UTF8);

            Console.WriteLine($"Formulas have been exported to '{csvOutputPath}'.");
        }
    }
}