// Title: Extract all formulas from an Excel workbook using Aspose.Cells for .NET and map them to a Dictionary with SheetName!CellAddress keys
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, iterates every worksheet and cell, detects formula cells, and adds each formula to a Dictionary<string,string> using "SheetName!CellAddress" as the key. | Enhance the extraction logic to also capture array and shared formulas, ensuring each unique formula is stored with its full address key in the same dictionary. | Add code that serializes the populated Dictionary of formulas to a JSON file using System.Text.Json after the extraction loop.
// Common Searches: how to get all cell formulas from an Excel file using Aspose.Cells in C# | Aspose.Cells C# iterate worksheets and collect formulas into a dictionary | store Excel formulas with sheet name prefix in a .NET dictionary using Aspose.Cells | extract array and shared formulas with Aspose.Cells .NET | serialize extracted Excel formulas to JSON using Aspose.Cells and System.Text.Json
// Tags: Aspose.Cells formula extraction | C# dictionary keyed by sheet name and cell address | iterate worksheets cells Aspose.Cells | read Excel cell formulas .NET | serialize extracted formulas to JSON

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, walks through each worksheet and every cell, checks for formulas, and stores each formula string in a Dictionary where the key combines the sheet name and cell address (e.g., Sheet1!A1). The collected formulas can then be printed or serialized to JSON.
class FormulaExtractor
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Dictionary to hold formulas keyed by full cell address (SheetName!CellAddress)
        Dictionary<string, string> formulas = new Dictionary<string, string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all cells in the current worksheet
            foreach (Cell cell in sheet.Cells)
            {
                // Check if the cell contains a formula
                if (cell.IsFormula)
                {
                    // Build a unique key using sheet name and cell address (e.g., Sheet1!A1)
                    string key = $"{sheet.Name}!{cell.Name}";

                    // Store the formula string in the dictionary
                    formulas[key] = cell.Formula;
                }
            }
        }

        // Example usage: print all extracted formulas
        foreach (var kvp in formulas)
        {
            Console.WriteLine($"{kvp.Key} = {kvp.Value}");
        }
    }
}
