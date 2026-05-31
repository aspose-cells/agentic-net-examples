using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace ExportFormulaCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path (change as needed)
            string inputPath = "input.xlsx";

            // Output CSV file path
            string outputPath = "formulas.csv";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure all formulas are calculated
            workbook.CalculateFormula();

            // Use the first worksheet (modify if another sheet is required)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Prepare CSV writer
            using (StreamWriter writer = new StreamWriter(outputPath, false, Encoding.UTF8))
            {
                // Write CSV header
                writer.WriteLine("Address,Formula,Result");

                // Iterate through all cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        // Cell address (e.g., A1)
                        string address = cell.Name;

                        // Formula string (as stored in the cell)
                        string formula = cell.Formula;

                        // Evaluated result of the formula
                        object resultObj = cell.Value;
                        string result = resultObj?.ToString() ?? string.Empty;

                        // Escape commas and double quotes in CSV fields
                        string Escape(string s)
                        {
                            if (s.Contains("\""))
                                s = s.Replace("\"", "\"\"");
                            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n"))
                                s = $"\"{s}\"";
                            return s;
                        }

                        // Write the CSV line
                        writer.WriteLine($"{Escape(address)},{Escape(formula)},{Escape(result)}");
                    }
                }
            }

            Console.WriteLine($"Formula export completed. CSV saved to: {outputPath}");
        }
    }
}