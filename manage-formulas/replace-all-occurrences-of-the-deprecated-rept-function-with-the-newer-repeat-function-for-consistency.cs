using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class ReplaceReptWithRepeatProgram
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through each cell that contains data
                    foreach (Cell cell in cells)
                    {
                        // Process only cells that have a formula
                        if (cell.IsFormula)
                        {
                            string originalFormula = cell.Formula;
                            string updatedFormula = ReplaceReptWithRepeatInFormula(originalFormula);

                            // Apply the new formula only if a change was made
                            if (!originalFormula.Equals(updatedFormula, StringComparison.Ordinal))
                            {
                                cell.Formula = updatedFormula;
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method that replaces all occurrences of "REPT(" (case‑insensitive) with "REPEAT("
        private static string ReplaceReptWithRepeatInFormula(string formula)
        {
            StringBuilder sb = new StringBuilder();
            int searchStart = 0;

            while (true)
            {
                int index = formula.IndexOf("REPT(", searchStart, StringComparison.OrdinalIgnoreCase);
                if (index == -1)
                {
                    sb.Append(formula.Substring(searchStart));
                    break;
                }

                sb.Append(formula.Substring(searchStart, index - searchStart));
                sb.Append("REPEAT(");
                searchStart = index + 5; // length of "REPT("
            }

            return sb.ToString();
        }
    }
}