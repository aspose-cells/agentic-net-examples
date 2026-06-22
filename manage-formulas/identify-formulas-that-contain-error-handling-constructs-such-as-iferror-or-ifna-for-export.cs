using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsErrorHandlingExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file (modify as needed)
            string inputPath = "SourceWorkbook.xlsx";

            // Path for the export file containing formulas with error handling
            string exportPath = "ErrorHandlingFormulas.csv";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(inputPath);

            // List to hold information about cells that contain IFERROR or IFNA
            List<string> exportLines = new List<string>();
            // Add CSV header
            exportLines.Add("SheetName,CellName,Formula");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate over all cells that have formulas
                foreach (Cell cell in cells)
                {
                    // Ensure the cell actually contains a formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Check for IFERROR or IFNA (case‑insensitive)
                        string formulaUpper = cell.Formula.ToUpperInvariant();
                        if (formulaUpper.Contains("IFERROR") || formulaUpper.Contains("IFNA"))
                        {
                            // Build a CSV line: SheetName,CellName,Formula
                            string line = $"{EscapeCsv(sheet.Name)},{EscapeCsv(cell.Name)},{EscapeCsv(cell.Formula)}";
                            exportLines.Add(line);
                        }
                    }
                }
            }

            // Write the collected data to the export CSV file
            File.WriteAllLines(exportPath, exportLines);

            // Optionally, save the workbook (unchanged) to demonstrate lifecycle usage
            workbook.Save("ProcessedWorkbook.xlsx");

            Console.WriteLine($"Export completed. {exportLines.Count - 1} formulas written to '{exportPath}'.");
        }

        // Helper method to escape CSV fields that may contain commas or quotes
        private static string EscapeCsv(string field)
        {
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
            {
                field = field.Replace("\"", "\"\"");
                return $"\"{field}\"";
            }
            return field;
        }
    }
}