// Title: Subclass Aspose.Cells AbstractCalculationEngine in C# to override Calculate and replace TODAY() with a fixed DATE literal
// Common Searches: C# Aspose.Cells replace TODAY() function with a specific date in workbook formulas | How to override Calculate in an Aspose.Cells custom calculation engine for fixed TODAY() value | Example of subclassing AbstractCalculationEngine to substitute TODAY() with DATE literal | Set a constant date for TODAY() before recalculating formulas using Aspose.Cells | Aspose.Cells custom engine to replace TODAY() with DATE(2023,1,1) in Excel files
// Tags: Aspose.Cells custom calculation engine | override Calculate method Aspose.Cells | replace TODAY function with DATE literal | fixed date substitution Excel formulas C# | recalculate workbook formulas after date replacement

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomEngine
{
    // Helper class to replace TODAY() with a fixed date in all formulas.
    // The example shows how to create a subclass of Aspose.Cells.AbstractCalculationEngine that overrides the Calculate method. It scans each cell, replaces any case‑insensitive TODAY() occurrences with a DATE(year,month,day) literal based on a supplied fixed DateTime, recalculates all formulas, and saves the updated workbook.
    public static class FixedDateHelper
    {
        // Replaces TODAY() occurrences with a DATE literal based on the provided fixed date.
        public static void ApplyFixedDate(Workbook workbook, DateTime fixedDate)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            string dateLiteral = $"DATE({fixedDate.Year},{fixedDate.Month},{fixedDate.Day})";

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    try
                    {
                        string formula = cell.Formula;
                        if (!string.IsNullOrEmpty(formula) &&
                            formula.IndexOf("TODAY()", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Replace all case‑insensitive occurrences of TODAY().
                            string newFormula = formula.Replace("TODAY()", dateLiteral, StringComparison.OrdinalIgnoreCase);
                            cell.Formula = newFormula;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log per‑cell errors without stopping the whole process.
                        Console.WriteLine($"Error processing cell {cell.Name}: {ex.Message}");
                    }
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to prevent FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook.
                Workbook workbook = new Workbook(inputPath);

                // Apply the fixed‑date substitution (e.g., 2023‑01‑01).
                FixedDateHelper.ApplyFixedDate(workbook, new DateTime(2023, 1, 1));

                // Recalculate all formulas after the substitution.
                workbook.CalculateFormula();

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors during processing.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
