// Title: C# method to clear Aspose.Cells calculation engine cache and force workbook formula recomputation
// Common Searches: C# clear Aspose.Cells calculation cache before recalculating formulas | reset formula cache in Aspose.Cells workbook programmatically | how to invalidate Aspose.Cells calculation engine cache after data changes | Aspose.Cells recalculate all formulas after clearing cache using C# | force fresh formula evaluation in Aspose.Cells by clearing cache
// Tags: Aspose.Cells calculation cache invalidation C# | Aspose.Cells formula cache reset | force formula recomputation Aspose.Cells | custom calculation engine clear Aspose.Cells | Workbook.CalculateFormula cache refresh

using System;
using System.IO;
using Aspose.Cells;

// Provides a static ClearCustomEngineCache method that validates the Workbook instance and invokes CalculateFormula inside a try‑catch block to discard any cached formula results. The example demonstrates loading a workbook, clearing the calculation engine cache, forcing a full formula recomputation, and saving the updated file.
public static class AsposeCellsHelper
{
    /// <param name="workbook">The Aspose.Cells Workbook instance.</param>
    public static void ClearCustomEngineCache(Workbook workbook)
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));

        try
        {
            // Recalculate all formulas to force the engine to discard previous cached values.
            workbook.CalculateFormula();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to clear calculation cache: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Determine the input file path (use a default if none provided)
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook from the specified file
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Clear the calculation engine cache
            AsposeCellsHelper.ClearCustomEngineCache(workbook);

            // Recalculate all formulas after clearing the cache
            workbook.CalculateFormula();

            // Prepare output path and ensure directory exists
            string outputDirectory = Path.GetDirectoryName(inputPath) ?? string.Empty;
            Directory.CreateDirectory(outputDirectory);
            string outputPath = Path.Combine(outputDirectory, "output.xlsx");

            // Save the processed workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
