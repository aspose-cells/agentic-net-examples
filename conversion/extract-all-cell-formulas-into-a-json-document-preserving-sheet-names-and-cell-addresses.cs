// Title: Export Excel formulas to a JSON file with sheet names and cell addresses – Aspose.Cells C# sample
// Description: This C# example uses Aspose.Cells to open an .xlsx workbook, walk through each worksheet, identify cells that contain formulas, and capture the worksheet name together with the A1‑style cell reference. The gathered information is then serialized into a formatted JSON array and written to disk.
// Keywords: Aspose.Cells | C# extract formulas | Excel to JSON | export formulas | worksheet name | cell address | .NET | JSON serialization | formula extraction | Aspose.Cells API | Excel automation
// Common Searches: how to export Excel formulas to JSON using Aspose.Cells | C# list all formula cells with sheet name | save workbook formulas as JSON file | Aspose.Cells get cell formula and address | convert Excel calculations to JSON .NET
// Developer Intent: Create a JSON document that enumerates every formula in an Excel workbook, including its sheet and cell reference.
// Use Cases: Document spreadsheet logic by exporting formulas and their locations for review or training. | Migrate calculation rules to another platform by providing a JSON map of sheet, address, and formula. | Perform regression testing on workbooks by comparing JSON snapshots of formulas before and after changes.
// AI Prompts: Generate C# code with Aspose.Cells that iterates all worksheets, collects cells where IsFormula is true, and writes a JSON array containing SheetName, Address, and Formula. | Show how to efficiently serialize a large list of formula objects to an indented JSON file using System.Text.Json in .NET. | Explain how to extend the sample to also include the evaluated value of each formula alongside the formula string.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// This C# example uses Aspose.Cells to open an .xlsx workbook, walk through each worksheet, identify cells that contain formulas, and capture the worksheet name together with the A1‑style cell reference. The gathered information is then serialized into a formatted JSON array and written to disk.
class ExtractFormulasToJson
{
    static void Main(string[] args)
    {
        // Path to the source Excel file
        string inputPath = "input.xlsx";

        // Path where the resulting JSON will be saved
        string outputPath = "formulas.json";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(inputPath);

        // Collection to hold information about each formula cell
        var formulaList = new List<FormulaInfo>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Iterate through all cells that are part of the used range
            foreach (Cell cell in cells)
            {
                // Check if the current cell contains a formula
                if (cell.IsFormula)
                {
                    formulaList.Add(new FormulaInfo
                    {
                        SheetName = sheet.Name,
                        Address   = cell.Name,   // e.g., "B6"
                        Formula   = cell.Formula // e.g., "=A1"
                    });
                }
            }
        }

        // Serialize the collection to a formatted JSON string
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        string json = JsonSerializer.Serialize(formulaList, jsonOptions);

        // Write the JSON string to the output file
        File.WriteAllText(outputPath, json);

        Console.WriteLine($"Extracted {formulaList.Count} formulas to '{outputPath}'.");
    }

    // Simple DTO representing a formula cell for JSON output
    private class FormulaInfo
    {
        public string SheetName { get; set; }
        public string Address   { get; set; }
        public string Formula   { get; set; }
    }
}
