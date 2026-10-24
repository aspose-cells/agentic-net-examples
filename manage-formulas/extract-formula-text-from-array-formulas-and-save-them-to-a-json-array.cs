// Title: C# code to extract unique array formulas from an Excel workbook and save them as a JSON array using Aspose.Cells
// Common Searches: C# Aspose.Cells read array formulas from Excel and export to JSON | How to get distinct array formula strings from a .xlsx file using Aspose.Cells | Serialize Excel array formulas to a JSON file in .NET | Extract and list unique array formulas with Aspose.Cells in C# | Save Excel array formula text to JSON using System.Text.Json
// Tags: extract array formulas Aspose.Cells | serialize formulas to JSON C# | iterate worksheets cells Aspose.Cells | collect unique array formulas .NET | load workbook without password Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet and its cells, identifies cells that belong to an array formula, gathers each distinct formula string, and then uses System.Text.Json to serialize the collection into a formatted JSON array written to a file. It includes error handling for workbook loading issues and general exceptions.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Path to the destination JSON file
        string jsonPath = "arrayFormulas.json";

        // Verify that the Excel file exists before attempting to load it
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
            return;
        }

        try
        {
            // Load the workbook (no password supplied). If the file is password‑protected,
            // a CellsException will be thrown and handled in the catch block below.
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(excelPath, loadOptions);

            // Collection to store unique array formula texts
            List<string> arrayFormulas = new List<string>();

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells that contain data
                foreach (Cell cell in cells)
                {
                    // Identify cells that are part of an array formula
                    if (cell.IsArrayFormula)
                    {
                        string formula = cell.Formula; // Formula text (includes leading '=')

                        // Add the formula if it hasn't been recorded yet
                        if (!arrayFormulas.Contains(formula))
                        {
                            arrayFormulas.Add(formula);
                        }
                    }
                }
            }

            // Convert the list of formulas to a formatted JSON string
            string json = JsonSerializer.Serialize(arrayFormulas, new JsonSerializerOptions { WriteIndented = true });

            // Write the JSON string to the output file
            File.WriteAllText(jsonPath, json);
            Console.WriteLine($"Array formulas have been saved to '{jsonPath}'.");
        }
        catch (CellsException ex)
        {
            // Handles errors related to loading the workbook (e.g., password protection)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
