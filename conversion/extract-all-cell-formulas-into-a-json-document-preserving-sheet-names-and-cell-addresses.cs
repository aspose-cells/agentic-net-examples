// Title: Extract all Excel cell formulas to a JSON file while preserving sheet names and A1 addresses using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells extract workbook formulas to JSON with sheet names | How to export Excel formulas as JSON using Aspose.Cells .NET | Save cell formulas with addresses to a JSON file in C# | Aspose.Cells iterate used range and retrieve formulas for each worksheet | Serialize extracted formulas to indented JSON using System.Text.Json | Extract formulas from a password‑protected Excel workbook with Aspose.Cells
// Tags: extract Excel formulas to JSON Aspose.Cells | Aspose.Cells get cell formula C# | serialize workbook formulas System.Text.Json | iterate used range cells Aspose.Cells | map sheet name to cell address Aspose.Cells | export workbook formulas without password Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace ExtractFormulasApp
{
    // The example loads an Excel workbook with Aspose.Cells, walks through each worksheet's used range, captures non‑empty formulas keyed by their A1 address, groups them by sheet name, and writes the resulting dictionary to a formatted JSON file using System.Text.Json, with error handling for Aspose.Cells‑specific and general exceptions.
    class ExtractFormulasToJson
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Path to the output JSON file
            string jsonPath = "formulas.json";

            // Verify that the input file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook (no password supplied)
                var loadOptions = new LoadOptions(); // default options
                Workbook workbook = new Workbook(excelPath, loadOptions);

                // Dictionary to hold sheet name -> (cell address -> formula) mappings
                var workbookFormulas = new Dictionary<string, Dictionary<string, string>>();

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Dictionary for the current sheet's formulas
                    var sheetFormulas = new Dictionary<string, string>();

                    // Get the used range of the sheet to limit iteration
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxColumn = cells.MaxDataColumn;

                    // Loop through all cells in the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxColumn; col++)
                        {
                            Cell cell = cells[row, col];
                            // Check if the cell contains a formula
                            if (!string.IsNullOrEmpty(cell.Formula))
                            {
                                // Use the cell's name (e.g., "A1") as the key
                                string address = cell.Name;
                                sheetFormulas[address] = cell.Formula;
                            }
                        }
                    }

                    // Add the sheet's formulas to the main dictionary if any formulas exist
                    if (sheetFormulas.Count > 0)
                    {
                        workbookFormulas[sheet.Name] = sheetFormulas;
                    }
                }

                // Serialize the dictionary to JSON with indentation for readability
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string jsonString = JsonSerializer.Serialize(workbookFormulas, jsonOptions);

                // Write the JSON string to the output file
                File.WriteAllText(jsonPath, jsonString);

                Console.WriteLine($"Formulas extracted to {jsonPath}");
            }
            catch (CellsException ex)
            {
                // Handles Aspose.Cells specific errors (e.g., password protection)
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handles any other unexpected errors
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
