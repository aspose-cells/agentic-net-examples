// Title: Extract Excel formulas into a C# Dictionary with Aspose.Cells
// Description: Loads a workbook, forces formula parsing, walks every worksheet and populated cell, checks the IsFormula flag, and stores each cell's A1 address together with its formula text in a Dictionary<string,string> that is returned to the caller.
// Keywords: Aspose.Cells formula extraction | C# Excel formula dictionary | parse formulas after load | .NET read Excel formulas | cell address to formula mapping | global | Aspose.Cells API
// Common Searches: Aspose.Cells get all formulas C# | dictionary of cell formulas Aspose | extract Excel formulas programmatically | iterate worksheets and collect formulas | force formula parsing Aspose.Cells
// Developer Intent: Collect every formula in a workbook and map it to its A1 cell reference.
// Use Cases: Generate an audit list of all formula cells for compliance reviews. | Batch‑replace or adjust formulas across multiple sheets before saving. | Export formula mappings to JSON, CSV, or a database for external analysis.
// AI Prompts: Create a C# method that opens an Excel file with Aspose.Cells, ensures formulas are parsed, and returns a Dictionary<string,string> of cell addresses and their formulas. | Show how to traverse all worksheets and cells in Aspose.Cells, selecting only those where IsFormula is true. | Explain how to handle workbooks loaded with formula parsing disabled and still retrieve the formulas using Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads a workbook, forces formula parsing, walks every worksheet and populated cell, checks the IsFormula flag, and stores each cell's A1 address together with its formula text in a Dictionary<string,string> that is returned to the caller.
class FormulaExtractor
{
    // Loads a workbook from the given path and extracts all formulas.
    // Returns a dictionary where the key is the cell address (e.g., "A1")
    // and the value is the formula string (e.g., "=SUM(B1:B5)").
    public static Dictionary<string, string> ExtractFormulas(string filePath)
    {
        // Load the workbook (uses default LoadOptions)
        Workbook workbook = new Workbook(filePath);

        // Ensure that any formulas that were not parsed on load are parsed now.
        // This avoids null or empty Formula values for cells that were loaded with parsing disabled.
        workbook.ParseFormulas(false);

        var formulas = new Dictionary<string, string>();

        // Iterate through each worksheet in the workbook.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Iterate through all cells that contain data in the worksheet.
            foreach (Cell cell in cells)
            {
                // Check if the cell actually contains a formula.
                if (cell.IsFormula)
                {
                    // cell.Name returns the address in A1 notation.
                    formulas[cell.Name] = cell.Formula;
                }
            }
        }

        return formulas;
    }

    // Example entry point demonstrating usage.
    static void Main()
    {
        string inputPath = "input.xlsx"; // Path to the workbook to process.

        Dictionary<string, string> formulaMap = ExtractFormulas(inputPath);

        // Output the extracted formulas.
        foreach (KeyValuePair<string, string> kvp in formulaMap)
        {
            Console.WriteLine($"{kvp.Key}: {kvp.Value}");
        }
    }
}
