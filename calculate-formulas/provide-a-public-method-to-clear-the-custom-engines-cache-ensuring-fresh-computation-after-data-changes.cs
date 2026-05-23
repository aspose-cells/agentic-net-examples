using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public static class CustomEngineCacheHelper
    {
        /// <summary>
        /// Clears any internal access caches that may affect a custom calculation engine.
        /// </summary>
        /// <param name="workbook">The workbook whose caches should be cleared.</param>
        public static void ClearCustomEngineCache(Workbook workbook)
        {
            // Close all active access caches.
            workbook.CloseAccessCache(AccessCacheOptions.All);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example file paths (adjust as needed)
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Prevent FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Clear custom engine caches
                CustomEngineCacheHelper.ClearCustomEngineCache(workbook);

                // Additional processing can be done here

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Runtime safety
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}