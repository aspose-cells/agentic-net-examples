// Title: Replace deprecated Excel functions in worksheet formulas with Aspose.Cells for .NET (C#)
// Common Searches: how to replace a deprecated Excel function in formulas using Aspose.Cells C# | Aspose.Cells C# iterate through all cells to update formula function names | search and replace old function name in workbook formulas with Aspose.Cells .NET | C# code to modify Excel formulas that contain a specific function using Aspose.Cells | replace function name in Excel formula across multiple worksheets with Aspose.Cells
// Tags: deprecated function substitution Aspose.Cells | update cell formulas C# | iterate worksheets modify formulas | regex formula name replacement Aspose.Cells | modify workbook formulas .NET

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an .xlsx workbook with Aspose.Cells, scans every worksheet and cell for formulas containing a specified deprecated function, replaces the function name with a new one using a case‑insensitive regular expression, and saves the updated workbook to a new file.
class Program
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
            // Load the workbook (no password)
            workbook = new Workbook(inputPath);
        }
        catch (CellsException ex) when (ex.Message.Contains("password"))
        {
            Console.WriteLine("The workbook is password protected. Unable to load without a password.");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Define deprecated function and its replacement
        string deprecatedFunction = "OLD_FUNC";
        string replacementFunction = "NEW_FUNC";

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain a formula
                    if (!string.IsNullOrEmpty(cell.Formula) &&
                        cell.Formula.IndexOf(deprecatedFunction, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace deprecated function name
                        string updatedFormula = ReplaceFunctionName(cell.Formula, deprecatedFunction, replacementFunction);
                        cell.Formula = updatedFormula;
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

    // Replaces the function name in a formula while preserving its arguments
    static string ReplaceFunctionName(string formula, string oldName, string newName)
    {
        string pattern = @"\b" + Regex.Escape(oldName) + @"\b";
        return Regex.Replace(formula, pattern, newName, RegexOptions.IgnoreCase);
    }
}
