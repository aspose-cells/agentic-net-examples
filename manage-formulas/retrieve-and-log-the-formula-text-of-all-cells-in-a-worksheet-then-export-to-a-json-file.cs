// Title: Retrieve formulas from every cell in an Excel worksheet and export them to a JSON file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get formula text for all cells in a worksheet | How to export Excel cell formulas to a JSON file using Aspose.Cells | Iterate used range and collect formulas with Aspose.Cells .NET | Save worksheet formulas as JSON with System.Text.Json and Aspose.Cells | Extract formula strings from Excel workbook using Aspose.Cells C#
// Tags: Aspose.Cells retrieve cell formulas | export formulas to JSON C# | iterate used range Aspose.Cells | Cell.IsFormula property Aspose.Cells | System.Text.Json serialize Excel data | extract formula text from worksheet

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The example loads a workbook, selects the first worksheet, iterates over its used range, checks each cell with the IsFormula property, gathers row, column, address, and formula text into a list, serializes the list to indented JSON using System.Text.Json, and writes the output to a file named formulas.json.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Select the worksheet you want to process (e.g., the first worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Prepare a collection to hold formula information
        List<object> formulaList = new List<object>();

        // Determine the used range of the worksheet
        Cells cells = worksheet.Cells;
        int maxRow = cells.MaxDataRow;
        int maxColumn = cells.MaxDataColumn;

        // Iterate through each cell in the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxColumn; col++)
            {
                Cell cell = cells[row, col];

                // Check if the cell contains a formula
                if (cell.IsFormula)
                {
                    // Capture row, column, address, and formula text
                    formulaList.Add(new
                    {
                        Row = row,
                        Column = col,
                        Address = cell.Name,
                        Formula = cell.Formula
                    });
                }
            }
        }

        // Convert the collection to JSON with indentation for readability
        string json = JsonSerializer.Serialize(formulaList, new JsonSerializerOptions { WriteIndented = true });

        // Write the JSON output to a file
        File.WriteAllText("formulas.json", json);
    }
}
