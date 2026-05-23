using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsFormulaExtractor
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Path where the JSON output will be saved
            string outputPath = "formulas.json";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Dictionary to hold sheet name -> list of formulas (address + formula)
            var workbookFormulas = new Dictionary<string, List<Dictionary<string, string>>>();

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var sheetFormulas = new List<Dictionary<string, string>>();
                Cells cells = sheet.Cells;

                // Iterate through all cells that contain data/formulas
                foreach (Cell cell in cells)
                {
                    // Check if the cell contains a formula
                    if (cell.IsFormula)
                    {
                        // Store the cell address and its formula
                        var formulaInfo = new Dictionary<string, string>
                        {
                            ["Address"] = cell.Name,   // e.g., "A1"
                            ["Formula"] = cell.Formula // e.g., "=SUM(B1:B5)"
                        };
                        sheetFormulas.Add(formulaInfo);
                    }
                }

                // Add to the result only if the sheet has at least one formula
                if (sheetFormulas.Count > 0)
                {
                    workbookFormulas[sheet.Name] = sheetFormulas;
                }
            }

            // Serialize the dictionary to a formatted JSON string
            string jsonOutput = JsonSerializer.Serialize(
                workbookFormulas,
                new JsonSerializerOptions { WriteIndented = true });

            // Write the JSON string to the output file
            File.WriteAllText(outputPath, jsonOutput);

            Console.WriteLine($"Formulas extracted and saved to '{outputPath}'.");
        }
    }
}