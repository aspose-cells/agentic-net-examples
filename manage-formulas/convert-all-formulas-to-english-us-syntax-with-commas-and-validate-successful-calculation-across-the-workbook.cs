// Title: Convert all Excel formulas to US English (comma) syntax and verify calculations with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change formula argument separator to commas for US English | How to recalculate all formulas after setting workbook culture to en-US in Aspose.Cells | Detect and log Excel formula errors using Aspose.Cells C# | Save an Excel workbook with updated US formula syntax using Aspose.Cells | Set workbook CultureInfo to en-US and enforce formula syntax with Aspose.Cells .NET
// Tags: convert Excel formulas to US comma syntax Aspose.Cells | recalculate workbook formulas after culture change C# | detect formula error values Aspose.Cells | save workbook with updated formula separators C# | set workbook culture en-US Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel file, sets the workbook's CultureInfo to en-US to enforce comma separators, recalculates all formulas, iterates through each cell to reassign formulas (ensuring US syntax) and logs any formula errors, then saves the modified workbook.
    class Program
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
            try
            {
                // Attempt to load the workbook (no password)
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

            try
            {
                // Set culture to en-US so formulas use commas as argument separators
                workbook.Settings.CultureInfo = new CultureInfo("en-US");

                // Recalculate all formulas with the new culture
                workbook.CalculateFormula();

                // Iterate through worksheets and cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    foreach (Cell cell in cells)
                    {
                        // Reassign formula to enforce US syntax
                        if (cell.IsFormula)
                        {
                            string currentFormula = cell.Formula;
                            cell.Formula = currentFormula;
                        }

                        // Detect formula errors
                        if (cell.IsFormula)
                        {
                            object value = cell.Value;
                            if (value is string error && error.StartsWith("#"))
                            {
                                Console.WriteLine($"Formula error in sheet '{sheet.Name}' cell {cell.Name}: {error}");
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}
