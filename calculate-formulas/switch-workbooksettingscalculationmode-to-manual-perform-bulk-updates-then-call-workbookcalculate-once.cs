// Title: Bulk insert cell values and trigger a single formula calculation in an Aspose.Cells workbook using C#
// Common Searches: Aspose.Cells C# bulk insert cell values and recalculate formulas once | How to disable automatic calculation in Aspose.Cells and run CalculateFormula manually | Set manual calculation mode for a workbook with Aspose.Cells .NET | Calculate all formulas after batch updates with Aspose.Cells in C# | Save an Aspose.Cells workbook to a specific folder after ensuring the directory exists | Performance tips for updating many cells in Aspose.Cells before a single calculation
// Tags: bulk cell value insertion Aspose.Cells | manual calculation setting Aspose.Cells | single formula recalculation Aspose.Cells | post‑batch formula calculation Aspose.Cells | ensure output directory before saving workbook | performance optimization Aspose.Cells bulk updates

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook (or loads one), fills a 1000 × 50 range with calculated values, then calls CalculateFormula once to evaluate all formulas, ensures the target folder exists, saves the file as UpdatedWorkbook.xlsx, and writes the saved path to the console.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // NOTE: Manual calculation mode is not set because the property may not be available
                // in some versions of Aspose.Cells. The workbook will be calculated explicitly later.

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Bulk update: fill a range with values
                for (int row = 0; row < 1000; row++)
                {
                    for (int col = 0; col < 50; col++)
                    {
                        sheet.Cells[row, col].PutValue(row * col);
                    }
                }

                // After all updates, trigger a single calculation
                workbook.CalculateFormula();

                // Save the workbook
                string outputPath = "UpdatedWorkbook.xlsx";

                // Ensure the directory exists (handle cases where outputPath has no directory part)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
