// Title: Replace deprecated CHOOSE function with nested IF formulas in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to replace CHOOSE with IF statements in Aspose.Cells C# | convert Excel CHOOSE formulas to nested IF using Aspose.Cells .NET | regex replace CHOOSE function in workbook cells with Aspose.Cells C# | migrate old CHOOSE formulas to IF logic in Excel using Aspose.Cells | split CHOOSE arguments and generate nested IF in C# with Aspose.Cells
// Tags: replace CHOOSE with nested IF Aspose.Cells | convert Excel CHOOSE formula .NET | regex based formula transformation Aspose.Cells | handle deprecated Excel functions C# | parse function arguments with nested parentheses C#

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an existing or new workbook, iterates through all cells, detects CHOOSE functions with a regular expression, converts each occurrence to an equivalent nested IF expression while correctly handling nested parentheses, updates the cell formulas, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                // LoadOptions can be used to specify a password if needed; here we assume none
                var loadOptions = new LoadOptions();
                workbook = new Workbook(inputPath, loadOptions);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one sheet
            }

            // Regex to locate CHOOSE functions: CHOOSE(index, value1, value2, ...)
            Regex chooseRegex = new Regex(@"CHOOSE\s*\(\s*([^,]+)\s*,\s*(.+?)\)", RegexOptions.IgnoreCase);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells that contain formulas
                foreach (Cell cell in sheet.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        string formula = cell.Formula;

                        // Replace each CHOOSE occurrence with nested IF statements
                        while (chooseRegex.IsMatch(formula))
                        {
                            formula = chooseRegex.Replace(formula, m => ConvertChooseToIf(m.Groups[1].Value, m.Groups[2].Value));
                        }

                        // Assign the transformed formula back to the cell
                        cell.Formula = formula;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Converts CHOOSE(index, val1, val2, ...) into IF(index=1,val1,IF(index=2,val2,...))
    static string ConvertChooseToIf(string indexPart, string valuesPart)
    {
        // Split the value arguments while respecting nested parentheses
        string[] values = SplitArguments(valuesPart);

        // Start with the last value as the default else part
        string nestedIf = values[values.Length - 1];

        // Build nested IFs from the end towards the beginning
        for (int i = values.Length - 2; i >= 0; i--)
        {
            // i+1 corresponds to the CHOOSE position (1‑based)
            nestedIf = $"IF({indexPart}={i + 1},{values[i]},{nestedIf})";
        }

        return nestedIf;
    }

    // Splits a comma‑separated argument list, handling nested parentheses
    static string[] SplitArguments(string args)
    {
        var list = new List<string>();
        int depth = 0;
        int start = 0;

        for (int i = 0; i < args.Length; i++)
        {
            char c = args[i];
            if (c == '(') depth++;
            else if (c == ')') depth--;
            else if (c == ',' && depth == 0)
            {
                list.Add(args.Substring(start, i - start).Trim());
                start = i + 1;
            }
        }

        // Add the final argument
        list.Add(args.Substring(start).Trim());

        return list.ToArray();
    }
}
