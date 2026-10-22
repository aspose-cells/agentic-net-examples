// Title: Programmatically replace semicolon argument separators with commas in Excel formulas and recalculate the workbook using Aspose.Cells for .NET
// Common Searches: how to change Excel formula delimiter from semicolon to comma using Aspose.Cells C# | Aspose.Cells C# replace semicolon in cell formulas and recalculate workbook | iterate through all cells to fix formula separators in .NET with Aspose.Cells | recalculate workbook after modifying formulas with Aspose.Cells API | convert European Excel formula separators to US style programmatically in C#
// Tags: Aspose.Cells replace formula delimiter | C# convert formula argument separator | recalculate all formulas Aspose.Cells | iterate worksheets cells modify formulas | Excel formula comma separator conversion

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, walks through every worksheet and cell, replaces any semicolon characters in cell formulas with commas, updates the formula when changed, forces a full recalculation of all formulas, and saves the corrected workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (no password). If the file is password‑protected,
            // a CellsException will be thrown and handled in the catch block.
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet and each cell to fix formulas
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        string correctedFormula = cell.Formula.Replace(";", ",");
                        if (correctedFormula != cell.Formula)
                        {
                            cell.Formula = correctedFormula;
                        }
                    }
                }
            }

            // Recalculate all formulas after modifications
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (CellsException ex)
        {
            // Handle password‑protected files or other Aspose.Cells specific errors
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // General exception handling
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
