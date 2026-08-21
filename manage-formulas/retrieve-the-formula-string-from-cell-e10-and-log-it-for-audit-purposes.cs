// Title: Extract and Log Cell E10 Formula (A1 Notation) with Aspose.Cells for .NET
// Description: Loads an Excel workbook, accesses cell E10 on the first worksheet, retrieves its formula in standard A1 notation using GetFormula(false, false), and writes the formula string to the console for auditing.
// Keywords: Aspose.Cells GetFormula | C# read Excel formula | extract cell formula .NET | audit Excel calculations | non‑localized A1 notation | log Excel formula | retrieve formula from cell E10
// Common Searches: Aspose.Cells read formula from specific cell C# | Get non‑localized formula using Aspose.Cells | How to log Excel cell formula for audit | Retrieve formula of E10 with Aspose.Cells .NET | Extract Excel formulas programmatically
// Developer Intent: Obtain the formula from cell E10 and output it for auditing.
// Use Cases: Verify that critical cells contain expected formulas before data processing. | Create an audit trail of spreadsheet logic by logging formulas from key cells. | Document worksheet calculations by extracting and storing formulas programmatically.
// AI Prompts: Generate C# code with Aspose.Cells that reads the formula from cell E10 and writes it to a log file. | Show how to loop through a range of cells, retrieve each formula, and save them in a dictionary using Aspose.Cells. | Explain the GetFormula parameters for obtaining localized versus non‑localized formulas in Aspose.Cells.

using Aspose.Cells;
using System;

// Loads an Excel workbook, accesses cell E10 on the first worksheet, retrieves its formula in standard A1 notation using GetFormula(false, false), and writes the formula string to the console for auditing.
class RetrieveFormula
{
    static void Main()
    {
        // Load the workbook containing the target cell
        string filePath = "input.xlsx"; // TODO: replace with actual file path
        Workbook workbook = new Workbook(filePath);

        // Access cell E10 on the first worksheet
        Cell cell = workbook.Worksheets[0].Cells["E10"];

        // Retrieve the formula in standard A1 notation (non‑localized)
        string formula = cell.GetFormula(false, false);

        // Log the retrieved formula for audit purposes
        Console.WriteLine($"Formula in E10: {formula}");
    }
}
