using System;
using Aspose.Cells;

namespace FormulaFixUtility
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // Input workbook path (replace with actual path)
            string inputPath = "input.xlsx";
            // Output workbook path
            string outputPath = "output_fixed.xlsx";

            FixMissingEqualSignFormulas(inputPath, outputPath);
            Console.WriteLine("Formula scan and fix completed.");
        }

        /// <summary>
        /// Scans all cells in the workbook and adds a leading '=' to strings that look like formulas.
        /// </summary>
        /// <param name="inputFile">Path to the source workbook.</param>
        /// <param name="outputFile">Path where the corrected workbook will be saved.</param>
        static void FixMissingEqualSignFormulas(string inputFile, string outputFile)
        {
            // Load the workbook (lifecycle rule: use provided load logic)
            Workbook workbook = new Workbook(inputFile);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to limit iteration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only string cells (potential missing formula)
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue;

                            // Skip empty strings and already correct formulas
                            if (string.IsNullOrWhiteSpace(text) || text.StartsWith("="))
                                continue;

                            // Simple heuristic: contains '(' and ')' suggests a formula
                            if (text.Contains("(") && text.Contains(")"))
                            {
                                // Prefix with '=' and assign to Formula property
                                cell.Formula = "=" + text;
                            }
                        }
                    }
                }
            }

            // Save the modified workbook (lifecycle rule: use provided save logic)
            workbook.Save(outputFile);
        }
    }
}