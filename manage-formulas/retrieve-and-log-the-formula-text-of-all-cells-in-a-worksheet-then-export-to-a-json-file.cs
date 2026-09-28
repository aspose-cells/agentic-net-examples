// Title: Extract formulas from an Excel worksheet and export them to a formatted JSON file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, scans the first worksheet for cells containing formulas, records each cell's address and formula string, and saves the collection as a pretty‑printed JSON document. | Update the sample program to include the worksheet name in each JSON entry and add a parameter to select which worksheet index to process. | Create an extension method for Aspose.Cells that returns a JSON string of all formula cells in a worksheet, with an option to skip cells located in hidden rows or columns.
// Common Searches: C# Aspose.Cells how to list all formula cells and export to JSON | Export Excel formulas to JSON file using Aspose.Cells .NET | Retrieve cell address and formula text from worksheet with Aspose.Cells C# | Save formulas from .xlsx to indented JSON using Aspose.Cells library
// Tags: aspocells extract formulas to json | c# read cell formula text aspocells | serialize worksheet formulas json c# | excel formula extraction aspocells .net | export formula cells as json aspocells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Loads an Excel workbook, iterates through every cell in the first worksheet, captures the address and formula of each formula cell, logs them to the console, and writes the collected data to an indented JSON file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "formulas.json";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Collection to store formula information
            List<FormulaInfo> formulaList = new List<FormulaInfo>();

            // Iterate through all cells in the worksheet
            foreach (Cell cell in worksheet.Cells)
            {
                // Check if the cell contains a formula
                if (cell.IsFormula)
                {
                    // Capture the cell address and its formula text
                    FormulaInfo info = new FormulaInfo
                    {
                        Address = cell.Name,
                        Formula = cell.Formula
                    };

                    // Add to the collection
                    formulaList.Add(info);

                    // Log to console
                    Console.WriteLine($"Cell {info.Address}: {info.Formula}");
                }
            }

            // Serialize the collection to JSON with indentation
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(formulaList, jsonOptions);

            // Write the JSON to a file
            File.WriteAllText(outputPath, json);
            Console.WriteLine($"Formulas have been written to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper class for JSON serialization
    class FormulaInfo
    {
        public string Address { get; set; }
        public string Formula { get; set; }
    }
}
