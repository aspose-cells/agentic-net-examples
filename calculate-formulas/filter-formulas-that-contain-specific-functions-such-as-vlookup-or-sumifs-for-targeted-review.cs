// Title: Use Aspose.Cells for .NET to highlight and list Excel cells that contain VLOOKUP or SUMIFS formulas
// AI Prompts: Iterate through every worksheet, detect formulas that include VLOOKUP or SUMIFS, apply a yellow background to those cells, and write their addresses to a new "FormulaReview" sheet using Aspose.Cells in C#. | Extend the filter to additional functions, change the highlight color to red, and export the matching cell references together with their original formulas to a CSV file with Aspose.Cells. | Modify the program to generate a summary workbook that groups identified formulas by worksheet and includes the full formula text beside each highlighted cell using Aspose.Cells.
// Common Searches: Aspose.Cells C# highlight cells containing VLOOKUP formula and create a review worksheet | list Excel cells with SUMIFS function using Aspose.Cells .NET | filter workbook formulas by function name and export results with Aspose.Cells
// Tags: highlight VLOOKUP and SUMIFS cells Aspose.Cells | collect formula addresses C# Aspose.Cells | search Excel workbook for function names .NET | generate formula review sheet Aspose.Cells | export matched formulas CSV Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

// Loads an Excel file, scans all worksheets for formulas containing VLOOKUP or SUMIFS, highlights matching cells, records their addresses on a new "FormulaReview" sheet, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // List to collect addresses of cells containing target functions
            List<string> targetFormulas = new List<string>();

            // Functions to search for
            string[] functions = new string[] { "VLOOKUP", "SUMIFS" };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range of the worksheet
                var usedRange = sheet.Cells.MaxDisplayRange;

                // Loop through each cell in the used range
                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;
                            foreach (string func in functions)
                            {
                                if (formula.IndexOf(func, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    // Record the cell address with sheet name
                                    targetFormulas.Add($"{sheet.Name}!{cell.Name}");

                                    // Highlight the cell for visual review
                                    Style style = cell.GetStyle();
                                    style.ForegroundColor = Color.Yellow;
                                    style.Pattern = BackgroundType.Solid;
                                    cell.SetStyle(style);

                                    break; // Move to next cell after first match
                                }
                            }
                        }
                    }
                }
            }

            // Create a new worksheet to list the identified formulas
            Worksheet resultSheet = workbook.Worksheets.Add("FormulaReview");
            resultSheet.Cells[0, 0].PutValue("Cell Address");

            for (int i = 0; i < targetFormulas.Count; i++)
            {
                resultSheet.Cells[i + 1, 0].PutValue(targetFormulas[i]);
            }

            // Save the workbook with the highlights and review sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Processing completed. Output saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
