// Title: C# – Retrieve and Log Formula from Cell E10 with Aspose.Cells
// Description: Load an Excel workbook, access the first worksheet, get the formula from cell E10 in standard A1 (non‑localized) format using GetFormula(false, false), write the formula to the console for audit, and optionally save the file.
// Keywords: Aspose.Cells GetFormula C# | read Excel cell formula .NET | audit Excel formulas | retrieve formula from E10 | log Excel formula C#
// Common Searches: Aspose.Cells read formula from specific cell | C# get formula of cell E10 Aspose | how to log Excel formula using Aspose.Cells | extract non‑localized formula string .NET
// Developer Intent: Extract the formula text from cell E10 and output it for auditing purposes.
// Use Cases: Create an audit trail of critical calculations in financial models. | Generate documentation that lists formulas used in key worksheet cells. | Validate that expected formulas are present before running automated processing.
// AI Prompts: Write a C# program that extracts formulas from a range of cells with Aspose.Cells and saves them to a CSV file. | Show how to compare a retrieved formula string with an expected pattern and throw an exception on mismatch. | Explain how to obtain formulas in R1C1 notation using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load an Excel workbook, access the first worksheet, get the formula from cell E10 in standard A1 (non‑localized) format using GetFormula(false, false), write the formula to the console for audit, and optionally save the file.
class RetrieveFormula
{
    static void Main()
    {
        // Load the workbook (replace with actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or specify by name/index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the cell at E10
        Cell cell = worksheet.Cells["E10"];

        // Get the formula string (standard A1 notation, non‑localized)
        string formula = cell.GetFormula(false, false);

        // Log the formula for audit purposes
        Console.WriteLine($"Formula in E10: {formula}");

        // Save the workbook (optional, if any changes were made)
        workbook.Save("output.xlsx");
    }
}
