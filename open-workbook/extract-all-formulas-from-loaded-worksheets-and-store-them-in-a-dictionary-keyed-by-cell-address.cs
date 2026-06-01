using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsFormulaExtractor
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Ensure all formulas are parsed (in case they were loaded without parsing)
            workbook.ParseFormulas(false);

            // Dictionary to hold formulas keyed by full cell address (e.g., Sheet1!A1)
            Dictionary<string, string> formulas = new Dictionary<string, string>();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells that contain data
                foreach (Cell cell in cells)
                {
                    // Check if the cell has a formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Build a unique key using sheet name and cell name (e.g., Sheet1!B2)
                        string key = $"{sheet.Name}!{cell.Name}";
                        formulas[key] = cell.Formula;
                    }
                }
            }

            // Example: output the collected formulas
            foreach (var kvp in formulas)
            {
                Console.WriteLine($"{kvp.Key} => {kvp.Value}");
            }

            // (Optional) Save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
    }
}