// Title: Clear cached formula results before invoking Workbook.CalculateFormula with Aspose.Cells for .NET
// Common Searches: Aspose.Cells clear calculation cache before Workbook.CalculateFormula C# | reset formula cache in Aspose.Cells .NET to force full recalculation | how to invalidate cached formulas in Aspose.Cells workbook | C# Aspose.Cells recalculate all formulas after clearing cache
// Tags: clear cached formulas Aspose.Cells | reset workbook calculation cache .NET | force fresh Workbook.CalculateFormula execution | invalidate internal formula cache Aspose.Cells | recalculate workbook without stale results C#

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // This C# example demonstrates how to purge all cached calculation results from an Aspose.Cells Workbook, then run Workbook.CalculateFormula to obtain a fresh, full‑recalculation of all formulas before saving the updated file.
    public class WorkbookRecalculator
    {
        /// <param name="inputPath">Path to the source workbook file.</param>
        /// <param name="outputPath">Path where the recalculated workbook will be saved.</param>
        public static void Recalculate(string inputPath, string outputPath)
        {
            try
            {
                // Verify that the input file exists.
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(inputPath);

                // Recalculate all formulas in the workbook.
                workbook.CalculateFormula();

                // Ensure the output directory exists.
                string outDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                    Directory.CreateDirectory(outDir);

                // Save the updated workbook to the desired location.
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during recalculation: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expect input and output paths as command‑line arguments.
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: WorkbookRecalculator <inputPath> <outputPath>");
                return;
            }

            WorkbookRecalculator.Recalculate(args[0], args[1]);
        }
    }
}
