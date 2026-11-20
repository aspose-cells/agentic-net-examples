// Title: Identify and highlight Excel cells that use VLOOKUP or SUMIFS formulas with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find cells containing VLOOKUP formula | filter workbook for SUMIFS functions using Aspose.Cells .NET | highlight cells with specific formulas in an .xlsx file via Aspose.Cells | list addresses of cells that contain VLOOKUP or SUMIFS in C# | apply background color to cells based on formula content with Aspose.Cells
// Tags: search formulas by function name Aspose.Cells | highlight cells based on formula content .NET | filter workbook formulas VLOOKUP SUMIFS | apply style to cells with specific Excel functions | extract cell addresses of target formulas C# | iterate worksheets used range Aspose.Cells | case‑insensitive formula search Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, scans each worksheet's used range, and collects cells whose formulas contain the functions VLOOKUP or SUMIFS (case‑insensitive). It prints the sheet name, address, and formula for each match, applies a light‑yellow background style to those cells, and saves the modified workbook to a new file.
class FormulaFilter
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "HighlightedWorkbook.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // List to hold cells with target formulas
            List<Cell> targetFormulaCells = new List<Cell>();

            // Functions to search for
            string[] functionsToFind = { "VLOOKUP", "SUMIFS" };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet (Aspose.Cells.Range)
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // Iterate through each cell in the used range
                foreach (Cell cell in usedRange)
                {
                    // Check if the cell contains a formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Determine if the formula contains any of the target functions
                        foreach (string func in functionsToFind)
                        {
                            if (cell.Formula.IndexOf(func, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                targetFormulaCells.Add(cell);
                                break; // No need to check other functions for this cell
                            }
                        }
                    }
                }
            }

            // Output the results
            Console.WriteLine("Cells containing specified functions:");
            foreach (Cell cell in targetFormulaCells)
            {
                Console.WriteLine($"Sheet: {cell.Worksheet.Name}, Address: {cell.Name}, Formula: {cell.Formula}");
            }

            // Highlight the found cells with a light yellow background
            foreach (Cell cell in targetFormulaCells)
            {
                Style style = cell.GetStyle();
                style.ForegroundColor = Color.LightYellow;
                style.Pattern = BackgroundType.Solid;
                cell.SetStyle(style);
            }

            // Save the highlighted workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Highlighted workbook saved as: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
