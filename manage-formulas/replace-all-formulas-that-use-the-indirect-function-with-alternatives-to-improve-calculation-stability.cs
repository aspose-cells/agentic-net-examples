// Title: C# example: replace INDIRECT formulas with direct cell references in an Excel workbook using Aspose.Cells
// Common Searches: how to replace INDIRECT function with direct cell reference in C# using Aspose.Cells | Aspose.Cells C# replace INDIRECT formulas for calculation stability | regex based formula modification in Excel workbooks with Aspose.Cells | load password protected Excel file and edit formulas using Aspose.Cells C# | convert dynamic INDIRECT references to static addresses with Aspose.Cells API
// Tags: replace indirect formulas Aspose.Cells C# | direct cell reference conversion Aspose.Cells | regex formula editing Aspose.Cells | password protected workbook loading Aspose.Cells | save modified workbook Aspose.Cells | excel calculation stability Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace ReplaceIndirectFormulasApp
{
    // The sample loads an Excel workbook (prompting for a password when required), iterates through all worksheets, detects cells that contain the INDIRECT function, uses a regular expression to replace the INDIRECT call with the actual cell address, updates the formula, and saves the modified workbook to a new file.
    class ReplaceIndirectFormulas
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook = null;

            // Attempt to load workbook; if password is required, prompt the user
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                // Simple detection of password protection based on exception message
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("The workbook is password protected. Please provide a password.");
                    Console.Write("Password: ");
                    string pwd = Console.ReadLine();

                    try
                    {
                        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = pwd };
                        workbook = new Workbook(inputPath, loadOptions);
                    }
                    catch (Exception innerEx)
                    {
                        Console.WriteLine($"Failed to load workbook with provided password: {innerEx.Message}");
                        return;
                    }
                }
                else
                {
                    Console.WriteLine($"Error loading workbook: {ex.Message}");
                    return;
                }
            }

            // Regex to match simple INDIRECT formulas
            Regex indirectRegex = new Regex(
                @"INDIRECT\(\s*(?:""(?<addr>[^""]+)""|(?<addr>[A-Za-z]+\d+))\s*\)",
                RegexOptions.IgnoreCase);

            // Iterate through worksheets and replace INDIRECT formulas
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula && cell.Formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string originalFormula = cell.Formula;
                        string newFormula = indirectRegex.Replace(originalFormula, m => m.Groups["addr"].Value);
                        if (!newFormula.Equals(originalFormula, StringComparison.Ordinal))
                        {
                            cell.Formula = newFormula;
                        }
                    }
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
