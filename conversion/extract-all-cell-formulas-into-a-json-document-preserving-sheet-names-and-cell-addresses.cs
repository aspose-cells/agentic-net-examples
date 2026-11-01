// Title: Extract all formulas from an Excel workbook to a JSON file with sheet names using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells extract cell formulas to JSON preserving worksheet names | How to export Excel formulas to a structured JSON document with Aspose.Cells | Save formulas from each sheet of an .xlsx as JSON using Aspose.Cells in .NET | Iterate used range and retrieve formula strings with Aspose.Cells C# example | Serialize Excel formula collection to indented JSON using System.Text.Json
// Tags: Aspose.Cells extract formulas to JSON | C# iterate worksheet used range | collect cell formula addresses | System.Text.Json serialize dictionary | group formulas by sheet name

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace ExtractFormulasToJson
{
    // The sample loads an .xlsx file with Aspose.Cells, walks through each worksheet's used range, captures cells that contain formulas along with their A1 addresses, groups the results by worksheet name in a dictionary, and writes the dictionary as formatted JSON using System.Text.Json, handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Path to the output JSON file
            string outputPath = "formulas.json";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook (using Aspose.Cells)
                Workbook workbook = new Workbook(inputPath);

                // Prepare a dictionary to hold sheet name -> list of cell formula info
                var workbookFormulas = new Dictionary<string, List<Dictionary<string, string>>>();

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // List to store formula details for the current sheet
                    var sheetFormulas = new List<Dictionary<string, string>>();

                    // Get the used range of the sheet to limit iteration to populated cells
                    var cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxColumn = cells.MaxDataColumn;

                    // Loop through each cell in the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxColumn; col++)
                        {
                            Cell cell = cells[row, col];

                            // Only process cells that contain a formula
                            if (!string.IsNullOrEmpty(cell.Formula))
                            {
                                // Build a dictionary with address and formula
                                var cellInfo = new Dictionary<string, string>
                                {
                                    { "Address", cell.Name },   // e.g., "A1"
                                    { "Formula", cell.Formula } // the formula string
                                };

                                sheetFormulas.Add(cellInfo);
                            }
                        }
                    }

                    // Add the sheet's formula list to the main dictionary (if any formulas were found)
                    if (sheetFormulas.Count > 0)
                    {
                        workbookFormulas[sheet.Name] = sheetFormulas;
                    }
                }

                // Serialize the dictionary to a formatted JSON string using System.Text.Json
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string jsonOutput = JsonSerializer.Serialize(workbookFormulas, jsonOptions);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write the JSON to the output file
                File.WriteAllText(outputPath, jsonOutput);

                Console.WriteLine($"Formulas extracted to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
