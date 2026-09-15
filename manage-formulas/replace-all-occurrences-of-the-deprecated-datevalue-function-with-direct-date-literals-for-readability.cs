// Title: How to replace deprecated DATEVALUE calls with DATE literals in Excel formulas using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, scans all formula cells, and swaps DATEVALUE("yyyy-mm-dd") with DATE(year,month,day) literals. | Write a method that uses a case‑insensitive regular expression to find DATEVALUE patterns in workbook formulas and rewrites them as DATE function calls. | Create a console application that processes an Excel workbook, converts each DATEVALUE formula to a DATE literal, and saves the result to a new file.
// Common Searches: Aspose.Cells replace DATEVALUE function with DATE literal in C# | C# regex to convert DATEVALUE("2023-01-01") to DATE(2023,1,1) in Excel formulas | How to update deprecated DATEVALUE formulas using Aspose.Cells workbook API | Programmatically change Excel date functions from DATEVALUE to DATE in .NET | Iterate through cells and modify formulas with Aspose.Cells C# example
// Tags: replace DATEVALUE with DATE literal Aspose.Cells | regex formula conversion Aspose.Cells C# | update Excel date functions programmatically .NET | convert DATEVALUE to DATE function Aspose.Cells | formula editing workbook cells C#

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The program loads an Excel workbook via Aspose.Cells, iterates through all used cells, detects DATEVALUE("yyyy-mm-dd") patterns in formulas using a case‑insensitive regex, parses the date string, replaces each occurrence with a DATE(year,month,day) literal, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Regex to find DATEVALUE("yyyy-mm-dd") patterns (case‑insensitive)
            Regex dateValueRegex = new Regex(@"DATEVALUE\(\s*""([^""]+)""\s*\)", RegexOptions.IgnoreCase);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Loop through all cells in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only formula cells
                        if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula))
                        {
                            string originalFormula = cell.Formula;

                            // Replace DATEVALUE("2023-01-01") with DATE(2023,1,1)
                            string updatedFormula = dateValueRegex.Replace(originalFormula, match =>
                            {
                                string dateText = match.Groups[1].Value;

                                // Try to parse the date string
                                if (DateTime.TryParse(dateText, out DateTime dt))
                                {
                                    // Build DATE(year, month, day) literal
                                    return $"DATE({dt.Year},{dt.Month},{dt.Day})";
                                }

                                // If parsing fails, keep the original DATEVALUE call
                                return match.Value;
                            });

                            // Apply the new formula if it changed
                            if (!updatedFormula.Equals(originalFormula, StringComparison.Ordinal))
                            {
                                cell.Formula = updatedFormula;
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
