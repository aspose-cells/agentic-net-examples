// Title: C# – Enable Aspose.Cells InconsistentFormula ErrorCheckOption and Retrieve Flagged Cells
// Description: Shows how to turn on ErrorCheckOption.InconsistentFormula in Aspose.Cells for .NET, run the calculation engine, and enumerate the cells that trigger the Inconsistent Formula warning.
// Keywords: Aspose.Cells | ErrorCheckOption | InconsistentFormula | .NET | C# | formula consistency | error checking | cell validation | Excel automation | detect inconsistent formulas
// Common Searches: Aspose.Cells enable InconsistentFormula check .NET | list cells flagged by InconsistentFormula warning | how to use ErrorCheckOption with Aspose.Cells | C# code to detect inconsistent formulas in Excel | retrieve error check results Aspose.Cells
// Developer Intent: Activate the InconsistentFormula error check and obtain the addresses of cells that raise the warning.
// Use Cases: Automated quality‑control that highlights formula patterns deviating from their neighbours. | Generating a diagnostic report of inconsistent formulas before workbook distribution. | Embedding formula‑consistency validation into CI pipelines that process generated Excel files.
// AI Prompts: Write C# code using Aspose.Cells to enable ErrorCheckOption.InconsistentFormula, run the check, and output the cell references that are flagged. | Show how to iterate over the ErrorCheckResult collection after activating InconsistentFormula checking and collect the cell names into a list. | Create a method that accepts a Workbook, turns on the InconsistentFormula warning, and returns an array of cell addresses with inconsistent formulas.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsErrorCheckDemo
{
    // Shows how to turn on ErrorCheckOption.InconsistentFormula in Aspose.Cells for .NET, run the calculation engine, and enumerate the cells that trigger the Inconsistent Formula warning.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate a region with formulas that are intentionally inconsistent
                cells["A1"].Formula = "=B1+1";
                cells["A2"].Formula = "=B2+1";
                cells["A3"].Formula = "=B3+2"; // Different formula – will be considered inconsistent

                cells["B1"].PutValue(10);
                cells["B2"].PutValue(20);
                cells["B3"].PutValue(30);

                // Calculate formulas so that values are updated
                workbook.CalculateFormula();

                // Output formulas and calculated values for the range A1:A3
                Console.WriteLine("Formulas and values in range A1:A3:");
                for (int row = 0; row < 3; row++)
                {
                    Cell cell = cells[row, 0]; // Column A
                    Console.WriteLine($"{cell.Name}: Formula = {cell.Formula}, Value = {cell.Value}");
                }

                // Save the workbook (optional, just to demonstrate lifecycle usage)
                string outputPath = "InconsistentFormulaCheckDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
