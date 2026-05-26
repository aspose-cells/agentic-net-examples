using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ReplaceDateValueFunction
    {
        // Converts a DATEVALUE formula to a DATE literal.
        private static string ConvertDateValueToDate(string formula)
        {
            const string target = "DATEVALUE(";
            int startIdx = formula.IndexOf(target, StringComparison.OrdinalIgnoreCase);
            if (startIdx == -1)
                return formula; // No DATEVALUE found

            int openParenIdx = startIdx + target.Length - 1;
            int closeParenIdx = formula.IndexOf(')', openParenIdx);
            if (closeParenIdx == -1)
                return formula; // Malformed formula

            string argument = formula.Substring(openParenIdx + 1, closeParenIdx - openParenIdx - 1).Trim();

            if (argument.StartsWith("\"") && argument.EndsWith("\""))
            {
                string dateString = argument.Substring(1, argument.Length - 2);
                if (DateTime.TryParse(dateString, out DateTime dt))
                {
                    string dateLiteral = $"DATE({dt.Year},{dt.Month},{dt.Day})";
                    return formula.Substring(0, startIdx) + dateLiteral + formula.Substring(closeParenIdx + 1);
                }
            }

            return formula;
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through worksheets and cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            if (cell.IsFormula && cell.Formula.IndexOf("DATEVALUE", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                string originalFormula = cell.Formula;
                                string updatedFormula = ConvertDateValueToDate(originalFormula);
                                if (!originalFormula.Equals(updatedFormula, StringComparison.Ordinal))
                                {
                                    cell.Formula = updatedFormula;
                                }
                            }
                        }
                    }
                }

                // Recalculate formulas after replacement (optional)
                workbook.CalculateFormula();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }

        // Entry point for the console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}