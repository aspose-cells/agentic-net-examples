// Title: Replace TODAY() formulas with a static date using Aspose.Cells for .NET (C#)
// Description: Loads a workbook, scans every worksheet and used cell, detects formulas that contain TODAY(), substitutes each with a supplied snapshot DateTime (e.g., 2023‑08‑15) via PutValue, and saves the modified file.
// Keywords: Aspose.Cells | C# | replace TODAY formula | static date | snapshot date | volatile function | Excel automation | cell iteration | Workbook.Save | date substitution
// Common Searches: Aspose.Cells replace TODAY() with fixed date | C# replace volatile TODAY function in Excel | How to set static date in Excel using Aspose.Cells | Iterate cells and modify formulas Aspose.Cells | Save workbook after changing formulas .NET
// Developer Intent: Replace every TODAY() formula in a workbook with a developer‑provided static date.
// Use Cases: Archive a workbook for audit trails by freezing all TODAY() calculations to a known snapshot date. | Generate financial reports that must retain the same reporting date across multiple distributions. | Batch‑process a set of Excel files to eliminate volatile date functions before publishing.
// AI Prompts: Write C# code with Aspose.Cells that scans all worksheets, replaces any TODAY() formula with a given DateTime, and saves the workbook. | Show how to log each cell address and its original formula while converting TODAY() to a static date using Aspose.Cells. | Explain how to apply culture‑specific date formatting when inserting a snapshot date with Aspose.Cells PutValue.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, scans every worksheet and used cell, detects formulas that contain TODAY(), substitutes each with a supplied snapshot DateTime (e.g., 2023‑08‑15) via PutValue, and saves the modified file.
    public class ReplaceTodayWithStaticDate
    {
        public static void Run()
        {
            try
            {
                // Path to the source workbook
                string inputPath = "input.xlsx";

                // Path to the resulting workbook
                string outputPath = "output.xlsx";

                // Snapshot date to replace TODAY() with (e.g., 2023‑08‑15)
                DateTime snapshotDate = new DateTime(2023, 8, 15);

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Loop through all used cells
                    foreach (Cell cell in cells)
                    {
                        // Check if the cell contains a formula that uses TODAY()
                        if (cell.IsFormula &&
                            !string.IsNullOrEmpty(cell.Formula) &&
                            cell.Formula.IndexOf("TODAY()", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Replace the formula with the static snapshot date
                            cell.PutValue(snapshotDate);
                        }
                    }
                }

                // Save the modified workbook (lifecycle rule: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ReplaceTodayWithStaticDate.Run();
        }
    }
}
