using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace ExtractArrayFormulas
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file (replace with your actual file path)
            string inputPath = "input.xlsx";

            // Load the workbook (lifecycle: load)
            Workbook workbook = new Workbook(inputPath);

            // List to hold the formula strings of all array formulas found
            List<string> arrayFormulas = new List<string>();

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to limit the iteration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Check if the cell contains an array formula
                        if (cell.IsArrayFormula)
                        {
                            // Add the formula text to the collection
                            arrayFormulas.Add(cell.Formula);
                        }
                    }
                }
            }

            // Serialize the list of formulas to a JSON array
            string json = JsonSerializer.Serialize(arrayFormulas, new JsonSerializerOptions { WriteIndented = true });

            // Path to the output JSON file
            string outputPath = "array_formulas.json";

            // Write the JSON content to the file (lifecycle: save)
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"Extracted {arrayFormulas.Count} array formula(s) and saved to '{outputPath}'.");
        }
    }
}