// Title: Export validated cells from the first worksheet of an Excel workbook to a JSON file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, iterates over the used range of the first worksheet, and writes each cell's address and value to a pretty‑printed JSON array on disk. | Create a method that verifies the input Excel file exists, ensures the output directory is created, and serializes a list of objects containing cell address and value to JSON using System.Text.Json. | Write robust error‑handling that logs cell‑level exceptions during iteration, continues processing remaining cells, and finally saves the collected data to a JSON file.
// Common Searches: how to use Aspose.Cells in C# to convert worksheet cells to JSON | C# export Excel used range to JSON array with Aspose.Cells | serialize cell address and value to JSON using Aspose.Cells .NET | Aspose.Cells example for writing validated cell data to a JSON file | handle missing Excel file and create output folder when exporting cells to JSON in C#
// Tags: Aspose.Cells export worksheet cells to JSON | C# serialize Excel cell address and value | used range iteration Aspose.Cells | System.Text.Json output file creation | cell processing error handling Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The sample loads an input.xlsx workbook, accesses the first worksheet, iterates through its used range, treats every cell as validated, collects each cell's address and value into a list, serializes the list to an indented JSON array with System.Text.Json, creates the output directory if necessary, writes the result to validated_cells.json, and logs any cell‑level errors while continuing processing.
class ExportValidatedCells
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "validated_cells.json";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet (avoid ambiguity with System.Range)
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startCol = usedRange.FirstColumn;
            int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            // List to hold validated cell information
            var validatedCells = new List<object>();

            // Iterate through each cell in the used range
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    try
                    {
                        Cell cell = sheet.Cells[row, col];

                        // In this example we treat all cells as "validated"
                        // Add cell address and value to the result list
                        validatedCells.Add(new
                        {
                            Address = cell.Name,
                            Value = cell.Value
                        });
                    }
                    catch (Exception cellEx)
                    {
                        // Log cell‑level errors but continue processing other cells
                        Console.WriteLine($"Cell processing error at {sheet.Cells[row, col].Name}: {cellEx.Message}");
                    }
                }
            }

            // Serialize the list of validated cells to JSON
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(validatedCells, jsonOptions);

            // Ensure the output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            catch (Exception dirEx)
            {
                Console.WriteLine($"Failed to create output directory: {dirEx.Message}");
            }

            // Write the JSON output to a file
            File.WriteAllText(outputPath, json);
            Console.WriteLine($"Validated cells have been exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
