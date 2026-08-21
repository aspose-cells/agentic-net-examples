// Title: Export All Worksheet Formulas to JSON with Aspose.Cells (C#/.NET)
// Description: Loads an Excel workbook using Aspose.Cells, optionally recalculates formulas, scans the used range of the first worksheet, captures each formula and its cell address, logs them, and writes the address‑formula pairs to a formatted JSON file (formulas.json).
// Keywords: Aspose.Cells C# extract formulas | export Excel formulas to JSON | list cell formulas Aspose.Cells | calculate workbook formulas .NET | serialize formulas JSON C#
// Common Searches: Aspose.Cells get formula text from cells | C# export all Excel formulas to JSON | iterate used range Aspose.Cells | save worksheet formulas as JSON file | recalculate formulas before export Aspose
// Developer Intent: Retrieve every formula in a worksheet, display it, and save the address‑formula pairs as JSON.
// Use Cases: Create an audit report of all formulas in a workbook. | Provide spreadsheet logic to a web API in JSON format. | Track changes to formulas after programmatic updates.
// AI Prompts: Write C# code that uses Aspose.Cells to load an Excel file, recalculate formulas, iterate the used range, and output each cell's address and formula to a JSON array. | Explain how to detect formula cells with Aspose.Cells and serialize their addresses and formulas into a pretty‑printed JSON file.

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// Loads an Excel workbook using Aspose.Cells, optionally recalculates formulas, scans the used range of the first worksheet, captures each formula and its cell address, logs them, and writes the address‑formula pairs to a formatted JSON file (formulas.json).
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Optional: calculate all formulas so that dependent values are up‑to‑date
        workbook.CalculateFormula();

        // Collect formula information
        var formulaList = new List<object>();

        // Determine the used range of the worksheet
        int maxRow = worksheet.Cells.MaxDataRow;
        int maxCol = worksheet.Cells.MaxDataColumn;

        // Iterate through each cell in the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = worksheet.Cells[row, col];
                if (cell.IsFormula)
                {
                    // Log formula to console
                    Console.WriteLine($"Cell {cell.Name}: {cell.Formula}");

                    // Store address and formula for JSON export
                    formulaList.Add(new
                    {
                        Address = cell.Name,
                        Formula = cell.Formula
                    });
                }
            }
        }

        // Serialize the collected formulas to a formatted JSON string
        string jsonOutput = JsonSerializer.Serialize(formulaList, new JsonSerializerOptions { WriteIndented = true });

        // Write the JSON string to a file
        File.WriteAllText("formulas.json", jsonOutput);
    }
}
