// Title: Replace Excel POWER function with ^ exponentiation operator using Aspose.Cells for .NET (C#)
// Common Searches: how to convert POWER(A1,2) to (A1)^(2) with Aspose.Cells C# | Aspose.Cells replace POWER formula with caret operator programmatically | C# regex to change Excel POWER function to exponentiation in a workbook | update all formulas in an Excel file from POWER to ^ using Aspose.Cells | load password‑protected workbook and modify POWER formulas with Aspose.Cells
// Tags: POWER to caret conversion Aspose.Cells | Excel exponentiation operator transformation | pattern based formula update Aspose.Cells | programmatic workbook formula modification .NET | load and edit password protected workbook Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook (prompting for a password when needed), iterates through every cell, uses a case‑insensitive regular expression to replace any POWER(a,b) formulas with (a)^(b), updates the cell formula only when a change occurs, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook without a password
            workbook = new Workbook(inputPath);
        }
        catch (CellsException ex) when (ex.Message.Contains("password", StringComparison.OrdinalIgnoreCase))
        {
            // Prompt for password if the workbook is protected
            Console.Write("Enter password for the workbook: ");
            string password = Console.ReadLine();

            try
            {
                var loadOptions = new LoadOptions { Password = password };
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook with provided password: {loadEx.Message}");
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
            // Iterate through all worksheets and replace POWER functions
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.IsFormula)
                    {
                        string originalFormula = cell.Formula;

                        // Replace POWER(A1,2) with (A1)^(2)
                        string updatedFormula = Regex.Replace(
                            originalFormula,
                            @"POWER\(\s*([^,]+?)\s*,\s*([^\)]+?)\s*\)",
                            "($1)^($2)",
                            RegexOptions.IgnoreCase);

                        if (!originalFormula.Equals(updatedFormula, StringComparison.Ordinal))
                        {
                            cell.Formula = updatedFormula;
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during processing: {ex.Message}");
        }
    }
}
