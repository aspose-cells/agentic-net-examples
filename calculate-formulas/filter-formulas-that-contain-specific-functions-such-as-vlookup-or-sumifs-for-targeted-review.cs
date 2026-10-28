// Title: Identify and highlight Excel cells with VLOOKUP or SUMIFS formulas using Aspose.Cells for .NET (C#)
// Common Searches: C# code to find cells that contain VLOOKUP formula in an .xlsx workbook with Aspose.Cells | how to highlight SUMIFS formulas in Excel using Aspose.Cells .NET library | list all formulas containing specific functions in a workbook with Aspose.Cells C# | filter Excel cells by formula function name using Aspose.Cells API | detect and mark VLOOKUP or SUMIFS formulas in a spreadsheet programmatically in C#
// Tags: formula function search Aspose.Cells | cell highlighting based on formula content C# | retrieve formula addresses for VLOOKUP SUMIFS | worksheet iteration for formula detection .NET | apply yellow background style to matched cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an .xlsx workbook, iterates through each worksheet and cell, detects formulas that contain the VLOOKUP or SUMIFS functions (case‑insensitive), records their addresses, highlights those cells with a yellow background, saves the modified file, and prints the matched formulas to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook (no password handling in this example)
            var loadOptions = new LoadOptions();
            workbook = new Workbook(inputPath, loadOptions);
        }
        catch (Exception ex)
        {
            // Simple detection of password‑protected files via message inspection
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The workbook is password protected. Unable to open without a password.");
            }
            else
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
            return;
        }

        // List to collect cells that contain the target functions
        List<string> matchedFormulas = new List<string>();

        // Functions to search for (case‑insensitive)
        string[] targetFunctions = { "VLOOKUP", "SUMIFS" };

        try
        {
            // Iterate through each worksheet in the workbook
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
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;

                            // Check if the formula contains any of the target functions
                            foreach (string func in targetFunctions)
                            {
                                if (formula.IndexOf(func, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    // Record the cell address and its formula
                                    matchedFormulas.Add($"{sheet.Name}!{cell.Name}: {formula}");

                                    // Highlight the cell for visual review
                                    Style style = cell.GetStyle();
                                    style.ForegroundColor = Color.Yellow;
                                    style.Pattern = BackgroundType.Solid;
                                    cell.SetStyle(style);

                                    // No need to check other functions once a match is found
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            // Save the workbook with highlighted cells
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
            return;
        }

        // Output the list of matched formulas to the console
        foreach (string entry in matchedFormulas)
        {
            Console.WriteLine(entry);
        }
    }
}
