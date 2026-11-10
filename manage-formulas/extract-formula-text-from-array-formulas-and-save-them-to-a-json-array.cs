// Title: Extract unique array formulas from an Excel workbook and save them as a JSON array using Aspose.Cells for .NET
// Common Searches: c# extract array formulas from Excel file using Aspose.Cells | how to get distinct array formula strings from a workbook with Aspose.Cells | save Excel array formulas to JSON file in .NET | Aspose.Cells iterate cells and collect IsArrayFormula values | remove duplicate array formulas when exporting from Excel in C#
// Tags: array formula extraction Aspose.Cells | distinct formula set C# | JSON serialization of formulas .NET | worksheet cell scanning Aspose.Cells | duplicate formula filtering HashSet

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsArrayFormulaExtractor
{
    // The program loads an Excel workbook with Aspose.Cells, scans each worksheet's used range, captures the Formula text of cells flagged as IsArrayFormula, stores them in a case‑insensitive HashSet to ensure uniqueness, then serializes the distinct formulas into a formatted JSON array and writes the result to a file, with error handling for missing files and I/O issues.
    class Program
    {
        static void Main()
        {
            // Paths for input workbook and output JSON
            string excelPath = "input.xlsx";
            string jsonPath = "arrayFormulas.json";

            // Verify that the source Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file '{excelPath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Use a set to avoid duplicate formulas from multiple cells in the same array range
                HashSet<string> formulaSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the used range limits
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    // Scan cells within the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Process only cells that belong to an array formula
                            if (cell.IsArrayFormula)
                            {
                                // Retrieve the formula text (includes leading '=')
                                // In some Aspose.Cells versions the array formula is accessed via Formula property
                                string formulaText = cell.Formula;
                                if (!string.IsNullOrEmpty(formulaText))
                                {
                                    formulaSet.Add(formulaText);
                                }
                            }
                        }
                    }
                }

                // Convert the set to a list for JSON serialization
                List<string> arrayFormulas = new List<string>(formulaSet);

                // Serialize to formatted JSON
                string json = JsonSerializer.Serialize(arrayFormulas, new JsonSerializerOptions { WriteIndented = true });

                // Ensure the directory for the JSON file exists
                try
                {
                    string directory = Path.GetDirectoryName(jsonPath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    // Write JSON to file (overwrites if exists)
                    File.WriteAllText(jsonPath, json);
                    Console.WriteLine($"Extracted {arrayFormulas.Count} unique array formulas and saved to '{jsonPath}'.");
                }
                catch (Exception ioEx)
                {
                    Console.WriteLine($"Failed to write JSON file: {ioEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }
}
