// Title: C# – Replace TODAY() formulas with a static snapshot date using Aspose.Cells
// Description: Loads an Excel file, scans every worksheet for cells that contain the TODAY() function, replaces each volatile formula with a fixed DateTime (e.g., 2023‑12‑31) via PutValue, and saves the updated workbook. Ideal for freezing dates in reports or archived files.
// Keywords: Aspose.Cells replace TODAY formula | static date for TODAY() | C# Excel volatile function removal | freeze TODAY() value | snapshot date Aspose.Cells
// Common Searches: how to replace TODAY() with a fixed date in Aspose.Cells | C# replace volatile Excel formulas Aspose.Cells | convert TODAY() to constant value .NET | freeze date functions in Excel using Aspose | Aspose.Cells replace TODAY() example
// Developer Intent: Replace every TODAY() formula in a workbook with a predefined static date.
// Use Cases: Generate reproducible reports where the reference date must not change between runs. | Archive financial workbooks for compliance by converting volatile date functions to constants. | Create test datasets with a known date baseline before sharing with stakeholders.
// AI Prompts: Write C# code that iterates through an Aspose.Cells workbook and substitutes any TODAY() formula with a given DateTime. | Show how to replace volatile functions like TODAY() and NOW() with static dates while keeping original cell formatting in Aspose.Cells. | Explain an efficient method to locate and update TODAY() formulas in large Excel files using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel file, scans every worksheet for cells that contain the TODAY() function, replaces each volatile formula with a fixed DateTime (e.g., 2023‑12‑31) via PutValue, and saves the updated workbook. Ideal for freezing dates in reports or archived files.
    public class ReplaceTodayWithStaticDate
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Core logic
        public static void Run()
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Path to the resulting workbook
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // The snapshot date that will replace all TODAY() formulas
            DateTime snapshotDate = new DateTime(2023, 12, 31);

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Loop through all used cells
                foreach (Cell cell in cells)
                {
                    // Process only formula cells containing TODAY()
                    if (cell.IsFormula &&
                        cell.Formula.IndexOf("TODAY()", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace the formula with the static snapshot date value
                        cell.PutValue(snapshotDate);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
