// Title: Export all formulas from an Excel workbook to a JSON file with Aspose.Cells for .NET (C#)
// Common Searches: how to extract all formulas from an .xlsx using Aspose.Cells in C# | C# code to write Excel cell formulas to a JSON document with Aspose.Cells | list worksheet formulas with sheet name and address using Aspose.Cells | serialize Aspose.Cells formula collection to indented JSON in .NET | export workbook formulas to JSON file handling missing file errors
// Tags: export formulas to JSON Aspose.Cells | extract cell formulas C# Aspose.Cells | list worksheet formulas with addresses | serialize Excel formulas using System.Text.Json | iterate workbook cells Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The example loads an input.xlsx workbook with Aspose.Cells, iterates through every worksheet and cell, gathers cells that contain formulas together with their sheet name and address, serializes the collection to a formatted JSON string using System.Text.Json, and writes the result to formulas.json while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "formulas.json";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Collect all formulas with their sheet name and address
            var formulaList = new List<object>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        formulaList.Add(new
                        {
                            SheetName = sheet.Name,
                            Address = cell.Name,
                            Formula = cell.Formula
                        });
                    }
                }
            }

            // Serialize the list to JSON with indentation
            var options = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(formulaList, options);

            // Write JSON to file
            File.WriteAllText(outputPath, json);
            Console.WriteLine($"Formulas have been exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
