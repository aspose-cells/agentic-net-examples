using System;
using Aspose.Cells;

namespace FormulaFilterDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Functions to look for
            string[] targetFunctions = { "VLOOKUP", "SUMIFS" };

            // Prepare find options: search only in formulas, allow partial match
            FindOptions options = new FindOptions
            {
                LookInType = LookInType.OnlyFormulas,
                LookAtType = LookAtType.Contains
            };

            // Iterate over each target function and locate matching cells
            foreach (string func in targetFunctions)
            {
                // Start search from the beginning each time
                Cell startCell = null;
                while (true)
                {
                    // Find the next cell containing the function name in its formula
                    Cell found = cells.Find(func, startCell, options);
                    if (found == null) break; // No more matches

                    // Output the address and the full formula
                    Console.WriteLine($"Found {func} in cell {found.Name}: {found.Formula}");

                    // Continue searching after the found cell
                    // Create a new start cell positioned after the current one
                    startCell = cells[found.Row, found.Column + 1];
                }
            }

            // Optionally, save the workbook (unchanged) to a new file
            workbook.Save("output.xlsx");
        }
    }
}