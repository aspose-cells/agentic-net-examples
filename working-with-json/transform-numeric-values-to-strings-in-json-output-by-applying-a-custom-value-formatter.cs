// Title: How to convert numeric Excel cells to quoted strings when saving as JSON using Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates through every worksheet and changes numeric cell values to their string representation before exporting to JSON with Aspose.Cells. | Show me how to apply a custom value formatter in Aspose.Cells so that numbers are written as quoted strings in the resulting JSON file. | Provide a step‑by‑step example that loads an .xlsx workbook, converts numeric cells to strings, and saves the workbook as JSON using JsonSaveOptions.
// Common Searches: Aspose.Cells .NET export Excel to JSON with numbers as strings | C# change numeric cell values to string before JSON serialization using Aspose.Cells | force quoted numeric values in JSON output from Excel workbook Aspose.Cells | how to use JsonSaveOptions to format numbers as strings in Aspose.Cells
// Tags: convert numeric cells to string Aspose.Cells | JSON export quoted numbers .NET | custom value formatter JsonSaveOptions Aspose.Cells | iterate worksheets cell conversion C# | Aspose.Cells numeric to string JSON output

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, walks through each worksheet and cell, replaces any numeric value with its string representation, and then saves the workbook as a JSON file using Aspose.Cells' JsonSaveOptions, ensuring numbers appear as quoted strings in the output.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.json";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Convert all numeric cell values to strings so they appear as quoted values in JSON
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        // Store the numeric value as its string representation
                        cell.PutValue(cell.Value.ToString());
                    }
                }
            }

            // Save the workbook as JSON using default options
            JsonSaveOptions jsonOptions = new JsonSaveOptions();
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook successfully saved as JSON to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
