// Title: Extract distinct array formula strings from an Excel workbook and export them to a JSON file using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, iterates every worksheet, identifies cells where IsArrayFormula is true, and collects each unique formula text into a List<string>. | Create a method that receives a List<string> of formulas and writes it as a pretty‑printed JSON array using System.Text.Json, ensuring the destination folder exists beforehand. | Add comprehensive error handling to the extraction routine to detect missing input files, catch I/O exceptions, and log the count of extracted array formulas.
// Common Searches: how to get all array formulas from an Excel file with Aspose.Cells in C# | C# Aspose.Cells extract unique formula strings and save as JSON | export Excel array formulas to JSON using .NET | iterate worksheets and check IsArrayFormula property Aspose.Cells example | handle missing workbook file when extracting formulas with Aspose.Cells
// Tags: Aspose.Cells extract array formulas | C# serialize formulas to JSON | unique Excel formula collection .NET | IsArrayFormula property usage | create output directory before file write

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The program loads an Excel workbook via Aspose.Cells, scans all worksheets and used cells, detects cells marked as array formulas (IsArrayFormula), gathers each distinct formula string, and writes the collection as an indented JSON array to a specified file while handling missing input files and ensuring the output directory exists.
class Program
{
    static void Main()
    {
        // Paths for input Excel and output JSON
        string inputPath = "input.xlsx";
        string outputPath = "arrayFormulas.json";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook (Aspose.Cells automatically detects the format)
            Workbook workbook = new Workbook(inputPath);

            // Collection for extracted array formula texts
            List<string> arrayFormulas = new List<string>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // If the worksheet is empty, skip it
                if (usedRange.RowCount == 0 || usedRange.ColumnCount == 0)
                    continue;

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Scan each cell within the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Identify array formulas
                        if (cell.IsArrayFormula)
                        {
                            string formulaText = cell.Formula; // includes leading '='

                            // Optional deduplication
                            if (!arrayFormulas.Contains(formulaText))
                            {
                                arrayFormulas.Add(formulaText);
                            }
                        }
                    }
                }
            }

            // Serialize the list to formatted JSON
            string json = JsonSerializer.Serialize(arrayFormulas, new JsonSerializerOptions { WriteIndented = true });

            // Ensure the output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write JSON to the output file
                File.WriteAllText(outputPath, json);
                Console.WriteLine($"Extracted {arrayFormulas.Count} array formulas and saved to '{outputPath}'.");
            }
            catch (Exception ioEx)
            {
                Console.WriteLine($"Failed to write output file: {ioEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
