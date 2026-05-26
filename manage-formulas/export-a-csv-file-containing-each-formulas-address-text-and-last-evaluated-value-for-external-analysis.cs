using System;
using System.IO;
using Aspose.Cells;

namespace ExportFormulasToCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path and output CSV file path
            string excelPath = "input.xlsx";
            string csvPath = "formulas.csv";

            // Create and load the workbook (lifecycle: create & load)
            Workbook workbook = new Workbook(excelPath);

            // Ensure all formulas are calculated so we can get the last evaluated values
            workbook.CalculateFormula();

            // Prepare a StreamWriter for the CSV output
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // Write CSV header
                writer.WriteLine("Address,Formula,Value");

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Determine the used range to limit iteration
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];

                            // Process only cells that contain a formula
                            if (!string.IsNullOrEmpty(cell.Formula))
                            {
                                string address = cell.Name;               // e.g., "A1"
                                string formula = cell.Formula;            // formula text
                                string value = cell.Value?.ToString() ?? ""; // last evaluated value

                                // Escape fields for CSV (wrap in quotes and double any internal quotes)
                                string csvAddress = EscapeCsv(address);
                                string csvFormula = EscapeCsv(formula);
                                string csvValue = EscapeCsv(value);

                                // Write the CSV line
                                writer.WriteLine($"{csvAddress},{csvFormula},{csvValue}");
                            }
                        }
                    }
                }
            }

            // At this point the CSV file is created (lifecycle: save handled by StreamWriter)
            Console.WriteLine($"Formulas exported to '{csvPath}'.");
        }

        // Helper method to escape a field for CSV format
        private static string EscapeCsv(string field)
        {
            if (field.Contains("\"") || field.Contains(",") || field.Contains("\n"))
            {
                field = field.Replace("\"", "\"\"");
                return $"\"{field}\"";
            }
            return field;
        }
    }
}