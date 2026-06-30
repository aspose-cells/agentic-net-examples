using System;
using System.Collections.Generic;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – extracts all formulas into a dictionary keyed by cell address
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Parse any formulas that were not parsed during load
        workbook.ParseFormulas(false);

        // Dictionary to store formulas: key = cell address (e.g., "A1"), value = formula string
        Dictionary<string, string> formulas = new Dictionary<string, string>();

        // Iterate through all worksheets and their cells
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            foreach (Cell cell in cells)
            {
                // If the cell contains a formula, add it to the dictionary
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    formulas[cell.Name] = cell.Formula;
                }
            }
        }

        // Example usage: display the extracted formulas
        foreach (var kvp in formulas)
        {
            Console.WriteLine($"{kvp.Key}: {kvp.Value}");
        }

        // Save the workbook if needed (unchanged in this example)
        workbook.Save("output.xlsx");
    }
}