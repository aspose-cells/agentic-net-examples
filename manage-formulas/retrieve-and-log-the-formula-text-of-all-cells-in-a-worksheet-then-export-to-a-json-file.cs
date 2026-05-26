using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // List to hold formula information
        var formulaList = new List<object>();

        // Determine the used range of the worksheet
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // Iterate through all cells in the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell != null && cell.IsFormula)
                {
                    // Log the cell address and its formula
                    Console.WriteLine($"{cell.Name}: {cell.Formula}");

                    // Store the formula data for JSON export
                    formulaList.Add(new
                    {
                        Address = cell.Name,
                        Formula = cell.Formula
                    });
                }
            }
        }

        // Serialize the collected formulas to a formatted JSON string
        string jsonOutput = JsonSerializer.Serialize(formulaList, new JsonSerializerOptions
        {
            WriteIndented = true
        });

        // Write the JSON string to a file
        File.WriteAllText("formulas.json", jsonOutput);
    }
}