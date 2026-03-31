using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsFormulaLister
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input XLSX workbook
            string inputPath = "input.xlsx";

            // Load the workbook (creation/loading rule)
            Workbook workbook = new Workbook(inputPath);

            // List to hold addresses of cells that contain formulas
            List<string> formulaCells = new List<string>();

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Check if the cell contains a formula
                    if (cell.IsFormula)
                    {
                        // Store the cell address in A1 notation with sheet name
                        formulaCells.Add($"{sheet.Name}!{cell.Name}");
                    }
                }
            }

            // Output the results
            Console.WriteLine("Cells containing formulas:");
            foreach (string address in formulaCells)
            {
                Console.WriteLine(address);
            }

            // (Optional) Save the workbook after any modifications (save rule)
            // workbook.Save("output.xlsx");
        }
    }
}