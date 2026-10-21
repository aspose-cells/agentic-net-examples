// Title: Replace RAND formulas with a seeded static value in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace volatile RAND formula with a constant value | How to make Excel RAND reproducible with a seeded random generator in Aspose.Cells | Iterate all worksheets and cells to convert RAND() to a fixed number using Aspose.Cells | Set a static seed for RAND() function in an Excel file with Aspose.Cells .NET | Replace random formulas with deterministic values in a workbook via Aspose.Cells
// Tags: replace RAND formula with static value Aspose.Cells | seeded random generator Excel .NET | iterate worksheets cells Aspose.Cells C# | convert volatile Excel formulas to constants | reproducible RAND function Aspose.Cells | fallback workbook creation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file (or creates a new one), initializes a seeded Random instance, walks through every worksheet and cell, detects formulas containing RAND, replaces each with a deterministic double from the seeded generator, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (CellsException ex)
                {
                    // If the file is password‑protected or corrupted, fall back to a new workbook.
                    Console.WriteLine($"Warning: Unable to load '{inputPath}'. Reason: {ex.Message}");
                    workbook = new Workbook();
                }
            }
            else
            {
                // No input file – create a fresh workbook with a sample RAND formula.
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sample";
                ws.Cells["A1"].Formula = "=RAND()";
            }

            // Seeded random generator for reproducible results.
            Random seededRandom = new Random(12345);

            // Iterate through all worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells in the worksheet.
                foreach (Cell cell in cells)
                {
                    // Process only cells that contain a RAND formula.
                    if (cell.IsFormula && cell.Formula.IndexOf("RAND", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        double staticValue = seededRandom.NextDouble();
                        cell.PutValue(staticValue); // Replace formula with static value.
                    }
                }
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // General exception handling to prevent crashes.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
