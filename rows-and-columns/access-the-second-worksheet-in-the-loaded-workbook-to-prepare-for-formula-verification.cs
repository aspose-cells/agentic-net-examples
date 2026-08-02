// Title: C# – Access the Second Worksheet and Verify Formulas with Aspose.Cells
// Description: Load an existing Excel file using Aspose.Cells, retrieve the worksheet at index 1 (the second sheet), run CalculateFormula to evaluate formulas, read a cell value to confirm the result, and save the workbook.
// Keywords: Aspose.Cells second worksheet C# | calculate formulas Aspose.Cells | Workbook.CalculateFormula example | read cell value after calculation | C# Excel worksheet index | Aspose.Cells formula verification
// Common Searches: Aspose.Cells get second sheet C# | how to calculate formulas on one worksheet Aspose.Cells | read cell after Workbook.CalculateFormula | C# Aspose.Cells verify formula results | save workbook after formula evaluation Aspose
// Developer Intent: Select the second sheet of a loaded workbook and confirm that its formulas are calculated correctly.
// Use Cases: Validate data integrity after bulk formula evaluation in a multi‑sheet workbook. | Improve performance by calculating formulas only on a targeted worksheet. | Automate post‑calculation checks such as reading specific cell values before saving.
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, accesses worksheet index 1, calculates formulas only on that sheet, and prints the value of cell B2. | Show how to handle formula errors returned by Workbook.CalculateFormula for a specific worksheet in Aspose.Cells. | Create a script that iterates through all worksheets, calculates formulas, and logs any cells that return errors.

using System;
using Aspose.Cells;

// Load an existing Excel file using Aspose.Cells, retrieve the worksheet at index 1 (the second sheet), run CalculateFormula to evaluate formulas, read a cell value to confirm the result, and save the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the second worksheet (index is zero‑based)
        Worksheet secondWorksheet = workbook.Worksheets[1];

        // Prepare for formula verification by calculating all formulas in the workbook
        // (or you could calculate only this sheet with workbook.CalculateFormula(secondWorksheet, ...))
        workbook.CalculateFormula();

        // Optionally, you can inspect a cell value to verify the calculation
        // Example: read the value of cell A1 in the second worksheet
        Cell cell = secondWorksheet.Cells["A1"];
        Console.WriteLine($"A1 value after calculation: {cell.Value}");

        // Save the workbook (if any changes were made)
        workbook.Save("output.xlsx");
    }
}
