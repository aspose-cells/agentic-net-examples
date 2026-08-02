// Title: Aspose.Cells for .NET – Enable Evaluation Error Checking and List #DIV/0! Cells
// Description: Demonstrates how to turn on evaluation‑error checking with ErrorCheckOptionCollection, calculate formulas, and programmatically collect the addresses of cells that return the #DIV/0! error in a workbook.
// Keywords: Aspose.Cells error checking | C# detect #DIV/0! cells | EvaluationError option Aspose | collect division by zero errors | Aspose.Cells formula validation
// Common Searches: Aspose.Cells enable evaluation error checking .NET | list cells with #DIV/0! using Aspose.Cells | ErrorCheckOptionCollection example C# | how to find division by zero errors in Aspose workbook | retrieve error cell addresses Aspose.Cells
// Developer Intent: Turn on formula evaluation error detection and extract the cell references that produce #DIV/0! after calculation.
// Use Cases: Generate a validation report highlighting all division‑by‑zero errors before publishing a workbook. | Automate data‑quality checks by flagging cells that contain #DIV/0! and applying corrective actions. | Integrate error detection into a larger ETL pipeline to prevent downstream processing failures.
// AI Prompts: Write C# code with Aspose.Cells that enables EvaluationError checking and returns a list of cell names containing #DIV/0! after workbook.CalculateFormula(). | Show how to configure ErrorCheckOptionCollection for a specific range to monitor evaluation errors and log offending cell addresses. | Provide an example that iterates through all cells, identifies any error value, and groups them by error type using Aspose.Cells for .NET.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsErrorCheckDemo
{
    // Demonstrates how to turn on evaluation‑error checking with ErrorCheckOptionCollection, calculate formulas, and programmatically collect the addresses of cells that return the #DIV/0! error in a workbook.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate some cells with formulas that will cause #DIV/0! errors
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(0);
            cells["B1"].Formula = "=A1/A2";          // Division by zero
            cells["B2"].Formula = "=A1/5";           // Valid formula
            cells["C1"].Formula = "=SUM(A1:A2)";     // Valid formula
            cells["C2"].Formula = "=A2/A2";          // 0/0 -> #DIV/0!

            // 3. Enable error checking for evaluation errors (e.g., #DIV/0!)
            ErrorCheckOptionCollection errorOptions = sheet.ErrorCheckOptions;
            int optionIdx = errorOptions.Add();                     // Add a new option
            ErrorCheckOption option = errorOptions[optionIdx];
            option.SetErrorCheck(ErrorCheckType.EvaluationError, true); // Enable checking
            // Apply the option to the whole used range
            CellArea usedArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = cells.MaxRow,
                EndColumn = cells.MaxColumn
            };
            option.AddRange(usedArea);

            // 4. Calculate all formulas (errors will be generated)
            workbook.CalculateFormula();

            // 5. Collect cells that contain the #DIV/0! error
            List<string> divZeroCells = new List<string>();
            foreach (Cell cell in cells)
            {
                if (cell.IsErrorValue && cell.StringValue == "#DIV/0!")
                {
                    divZeroCells.Add(cell.Name);
                }
            }

            // 6. Output the addresses of cells with #DIV/0! errors
            Console.WriteLine("Cells containing #DIV/0! error:");
            foreach (string address in divZeroCells)
            {
                Console.WriteLine(address);
            }

            // 7. Save the workbook (optional, demonstrates lifecycle rule)
            workbook.Save("ErrorCheckDivZeroDemo.xlsx");
        }
    }
}
