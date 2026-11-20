// Title: Use Aspose.Cells in C# to calculate formulas in parallel for large Excel workbooks
// Common Searches: C# Aspose.Cells parallel formula calculation for large Excel workbooks | how to speed up Workbook.CalculateFormula using multiple cores in Aspose.Cells | replace default calculation engine with a parallel version in Aspose.Cells .NET | enable multi‑threaded formula recalculation for big spreadsheets with Aspose.Cells | optimize performance of Excel workbook recalculation using Aspose.Cells C#
// Tags: Aspose.Cells multi‑core calculation engine | parallel workbook formula processing .NET | large spreadsheet performance optimization C# | custom calculation plugin for Aspose.Cells | accelerated formula evaluation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsParallelPlugin
{
    // The example shows a C# plugin that loads an Excel workbook, calls Workbook.CalculateFormula() which automatically leverages Aspose.Cells' built‑in parallel calculation when the workbook size and environment permit, and then saves the workbook. It includes error handling for missing files and ensures the output directory exists, providing a simple console entry point for integration.
    public class ParallelCalculationPlugin
    {
        public void Process(string inputPath, string outputPath)
        {
            try
            {
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(inputPath);

                // Recalculate all formulas. Aspose.Cells will use parallel calculation
                // automatically when the environment and workbook size allow it.
                workbook.CalculateFormula();

                // Ensure the output directory exists.
                string outDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                    Directory.CreateDirectory(outDir);

                // Save the processed workbook.
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
                throw;
            }
        }
    }

    // Simple console entry point for testing the plugin.
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsParallelPlugin <inputPath> <outputPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            var plugin = new ParallelCalculationPlugin();
            plugin.Process(inputPath, outputPath);
        }
    }
}
