// Title: How to replace DATEVALUE formulas with literal dates in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace DATEVALUE("2023-01-01") formula with actual date value | Convert DATEVALUE function to a date literal in Excel using Aspose.Cells .NET | Remove DATEVALUE formulas from all sheets in a password‑protected workbook with Aspose.Cells | Update Excel formulas containing DATEVALUE to DateTime values via Aspose.Cells API | Regex replace DATEVALUE in workbook formulas using Aspose.Cells C# example
// Tags: replace DATEVALUE formula Aspose.Cells | convert DATEVALUE to date literal .NET | regex based formula modification Aspose.Cells | load password protected workbook Aspose.Cells | update Excel formulas across worksheets C#

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace ReplaceDateValueApp
{
    // The C# program loads an Excel workbook (prompting for a password when required), scans every worksheet for formulas that use the DATEVALUE("…") pattern, parses the extracted date strings into DateTime objects, replaces the formulas with the corresponding literal date values, and saves the updated workbook to a new file.
    class ReplaceDateValue
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook = null;

            try
            {
                // Attempt to load the workbook without a password
                workbook = new Workbook(inputPath);
            }
            catch (CellsException)
            {
                // If a password is required, ask the user and retry loading
                Console.WriteLine("The workbook is password protected.");
                Console.Write("Enter password: ");
                string password = Console.ReadLine();

                var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
                try
                {
                    workbook = new Workbook(inputPath, loadOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook with password: {ex.Message}");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            try
            {
                // Regex to match DATEVALUE("yyyy-mm-dd") patterns
                Regex dateValueRegex = new Regex(@"DATEVALUE\(\s*""([^""]+)""\s*\)", RegexOptions.IgnoreCase);

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    var cells = sheet.Cells;
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
                                Match match = dateValueRegex.Match(cell.Formula);
                                if (match.Success)
                                {
                                    string dateString = match.Groups[1].Value;

                                    // Convert the extracted string to a DateTime and replace the formula
                                    if (DateTime.TryParse(dateString, out DateTime parsedDate))
                                    {
                                        cell.PutValue(parsedDate);
                                    }
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
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
