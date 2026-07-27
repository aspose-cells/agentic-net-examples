// Title: C# – Extract every formula from an Excel workbook into a SheetName!Cell dictionary with Aspose.Cells
// Description: Loads an XLSX file using Aspose.Cells, forces formula parsing, scans each worksheet’s used range, captures non‑empty Formula values, and stores them in a Dictionary where the key is "SheetName!CellAddress". The dictionary can be printed, saved, or further processed.
// Keywords: Aspose.Cells extract formulas C# | C# get all Excel formulas dictionary | LoadOptions ParsingFormulaOnOpen | Workbook.ParseFormulas | iterate worksheets Aspose.Cells | cell.Formula property | .NET Excel audit | global developers | US .NET community | European data processing
// Common Searches: how to list all formulas in an Excel file using Aspose.Cells C# | dictionary of sheet name and cell address for formulas .NET | parse formulas after opening workbook Aspose.Cells | retrieve cell.Formula for every cell in workbook | extract Excel formulas for audit report C#
// Developer Intent: Gather every formula present in a loaded workbook and organize them in a key‑value map keyed by sheet name and cell address.
// Use Cases: Create a formula audit log that shows which cells contain calculations. | Replace or update specific formulas programmatically after locating them in the dictionary. | Export the formula map to JSON, CSV, or a database for external analysis or reporting.
// AI Prompts: Generate a reusable C# method that returns Dictionary<string,string> of all formulas from a Workbook using Aspose.Cells. | Provide C# code that extracts formulas while correctly handling merged cells, hidden rows, and protected sheets. | Show how to serialize the formula dictionary to JSON and write it to a file with System.Text.Json.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an XLSX file using Aspose.Cells, forces formula parsing, scans each worksheet’s used range, captures non‑empty Formula values, and stores them in a Dictionary where the key is "SheetName!CellAddress". The dictionary can be printed, saved, or further processed.
class ExtractFormulas
{
    static void Main()
    {
        // Load the workbook (adjust the file path as needed)
        string inputPath = "input.xlsx";

        // Ensure formulas are parsed when the file is opened
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = true;
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Parse any formulas that were not parsed during load
        workbook.ParseFormulas(false);

        // Dictionary to store formulas keyed by "SheetName!CellAddress"
        Dictionary<string, string> formulaDictionary = new Dictionary<string, string>();

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Scan each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // If the cell contains a formula, add it to the dictionary
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        string key = $"{sheet.Name}!{cell.Name}";
                        formulaDictionary[key] = cell.Formula;
                    }
                }
            }
        }

        // Output the collected formulas (optional)
        foreach (var kvp in formulaDictionary)
        {
            Console.WriteLine($"{kvp.Key} => {kvp.Value}");
        }

        // Save the workbook if any changes were made (optional)
        // workbook.Save("output.xlsx");
    }
}
