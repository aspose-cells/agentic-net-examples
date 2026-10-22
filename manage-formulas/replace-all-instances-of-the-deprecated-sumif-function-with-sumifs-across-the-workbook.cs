// Title: Replace deprecated SUMIF formulas with SUMIFS across all worksheets using Aspose.Cells for .NET (C#)
// Common Searches: C# code to update Excel workbook formulas from SUMIF to SUMIFS with Aspose.Cells | How to programmatically replace SUMIF with SUMIFS in every sheet using Aspose.Cells .NET | Iterate through all cells in a workbook and modify formulas using Aspose.Cells C# | Using regular expressions to change Excel functions in Aspose.Cells | Bulk formula conversion from SUMIF to SUMIFS in an XLSX file with Aspose.Cells | Aspose.Cells replace deprecated Excel functions in a .NET application
// Tags: SUMIF to SUMIFS conversion Aspose.Cells | bulk Excel formula replacement .NET | worksheet iteration for formula editing C# | regex based formula modification Aspose.Cells | deprecated Excel function handling .NET

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace WorkbookFormulaUpdater
{
    // The example loads an Excel workbook (creating a placeholder if missing), iterates through every worksheet and cell, detects formulas containing the SUMIF function, replaces each occurrence with SUMIFS using a case‑insensitive regular expression, and saves the updated workbook while handling Aspose.Cells‑specific and general exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a placeholder if it does not.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook as a placeholder.");
                    var placeholder = new Workbook();
                    placeholder.Worksheets[0].Name = "Sheet1";
                    placeholder.Save(inputPath);
                }

                // Load the workbook (no password required).
                var loadOptions = new LoadOptions();
                var workbook = new Workbook(inputPath, loadOptions);

                // Iterate through all worksheets and replace SUMIF with SUMIFS in formulas.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    foreach (Cell cell in cells)
                    {
                        if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula))
                        {
                            if (cell.Formula.IndexOf("SUMIF(", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                string newFormula = Regex.Replace(
                                    cell.Formula,
                                    @"\bSUMIF\s*\(",
                                    "SUMIFS(",
                                    RegexOptions.IgnoreCase);
                                cell.Formula = newFormula;
                            }
                        }
                    }
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
