// Title: C# extension method to clear Aspose.Cells custom calculation engine cache by forcing full workbook recalculation
// AI Prompts: Generate a C# Workbook extension named ClearCustomEngineCache that forces a full recalculation to purge cached formula values using Aspose.Cells. | Show how to load an Excel file, invoke ClearCustomEngineCache on the Workbook, and save the result. | Explain how the extension handles a null Workbook argument and why CalculateFormula(true) clears the cache.
// Common Searches: how to purge formula cache in Aspose.Cells C# after modifying cell values | Aspose.Cells C# force full workbook recalculation to reset cached results | extension method to clear custom calculation engine cache in Aspose.Cells | reset Aspose.Cells calculation cache without reopening the workbook | C# clear cached formulas in an existing Excel workbook using Aspose.Cells
// Tags: Aspose.Cells workbook cache reset using CalculateFormula | C# extension method clear calculation cache | Aspose.Cells full recalculation API .NET | custom engine cache invalidation Aspose.Cells | reset Excel formula cache after data change

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Adds a static WorkbookExtensions class with a ClearCustomEngineCache extension method that calls CalculateFormula(true) to force a full recalculation, thereby clearing cached formula results. Demonstrates loading a workbook, invoking the method, and saving the updated file.
    public static class WorkbookExtensions
    {
        /// <param name="workbook">The workbook whose calculation cache should be cleared.</param>
        public static void ClearCustomEngineCache(this Workbook workbook)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            // Force full recalculation which clears cached values.
            workbook.CalculateFormula(true);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "sample.xlsx";

                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook.
                Workbook wb = new Workbook(inputPath);

                // Clear calculation cache.
                wb.ClearCustomEngineCache();

                // Save the processed workbook.
                string outputPath = "output.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook processed and saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
