// Title: C# – Load a workbook and set a SUM formula in cell B2 with Aspose.Cells
// Description: Demonstrates how to open an existing Excel file using Aspose.Cells for .NET, assign the formula =SUM(A1:A5) to cell B2, trigger calculation, output the result, and save the workbook as a new file.
// Keywords: Aspose.Cells | C# | set cell formula | SUM formula | load workbook | calculate formulas | save Excel file | Excel automation .NET | cell B2 formula | Workbook.CalculateFormula
// Common Searches: Aspose.Cells set formula C# | How to add SUM to a cell with Aspose.Cells | Recalculate workbook after formula Aspose.Cells .NET | Save Excel after modifying formulas C# | Assign formula to specific cell using Aspose.Cells
// Developer Intent: The programmer needs to open an existing Excel workbook, programmatically place a SUM expression in cell B2, evaluate the sheet, and write the changes back to disk.
// Use Cases: Generate a totals row in financial spreadsheets automatically | Prepare summary calculations in a template before distribution | Refresh derived values after bulk data import in an export file
// AI Prompts: Write C# code that uses Aspose.Cells to insert a custom formula into cell D5, recalculate the sheet, and save the file. | Show how to loop through a range of cells and assign different formulas with error handling in Aspose.Cells. | Explain how to check whether a formula result is numeric before saving the workbook using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    // Demonstrates how to open an existing Excel file using Aspose.Cells for .NET, assign the formula =SUM(A1:A5) to cell B2, trigger calculation, output the result, and save the workbook as a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook from file
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the target cell B2
            Cell targetCell = worksheet.Cells["B2"];

            // Assign a SUM formula to B2 (e.g., sum of A1:A5)
            targetCell.Formula = "=SUM(A1:A5)";

            // Calculate the workbook so the formula result is evaluated
            workbook.CalculateFormula();

            // Optionally display the calculated value
            Console.WriteLine("B2 value after calculation: " + targetCell.Value);

            // Save the modified workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
