// Title: Extract all formulas from an Excel workbook into a SheetName!CellAddress dictionary with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells extract formulas from every worksheet into a dictionary | How to get a list of cell formulas with sheet name as key using Aspose.Cells | Iterate used range of Excel file and collect formulas with Aspose.Cells .NET | Retrieve formula strings from Excel workbook programmatically in C#
// Tags: extract formulas Aspose.Cells C# | dictionary keyed by sheet name and cell address | iterate worksheets used range Aspose.Cells | cell.IsFormula property usage | retrieve formula string Excel .NET

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example shows how to load an Excel file with Aspose.Cells, loop through each worksheet’s used rows and columns, detect cells that contain formulas via Cell.IsFormula, build a key in the form "SheetName!A1", store the formula text (Cell.Formula) in a Dictionary<string,string>, and return the collection. The Main method demonstrates calling the method and printing the results.
class FormulaExtractor
{
    // Extracts all formulas from all worksheets in the given workbook.
    // Returns a dictionary where the key is "SheetName!CellAddress" (e.g., "Sheet1!A1")
    // and the value is the formula string (e.g., "=SUM(B1:B10)").
    public static Dictionary<string, string> ExtractFormulas(string workbookPath)
    {
        // Load the workbook from the specified file.
        Workbook workbook = new Workbook(workbookPath);

        // Dictionary to hold the formulas.
        Dictionary<string, string> formulas = new Dictionary<string, string>();

        // Iterate through each worksheet in the workbook.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range of the sheet.
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Loop through all cells within the used range.
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Check if the cell contains a formula.
                    if (cell.IsFormula)
                    {
                        // Build a unique key that includes the sheet name.
                        string key = $"{sheet.Name}!{cell.Name}";

                        // Store the formula text.
                        formulas[key] = cell.Formula;
                    }
                }
            }
        }

        return formulas;
    }

    // Example usage.
    static void Main()
    {
        string path = "input.xlsx"; // Replace with your workbook file path.
        Dictionary<string, string> allFormulas = ExtractFormulas(path);

        // Output the extracted formulas.
        foreach (var kvp in allFormulas)
        {
            Console.WriteLine($"{kvp.Key} => {kvp.Value}");
        }
    }
}
