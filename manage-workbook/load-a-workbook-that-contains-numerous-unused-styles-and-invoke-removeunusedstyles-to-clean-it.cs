using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveUnusedStylesDemo
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Path to the workbook that contains many unused styles
                string inputPath = "InputWithUnusedStyles.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the file path
                Workbook wb = new Workbook(inputPath);

                // Show the number of styles before removal
                Console.WriteLine($"Styles before cleaning: {wb.CountOfStylesInPool}");

                // Remove all unused styles
                wb.RemoveUnusedStyles();

                // Show the number of styles after removal
                Console.WriteLine($"Styles after cleaning: {wb.CountOfStylesInPool}");

                // Save the cleaned workbook
                string outputPath = "CleanedWorkbook.xlsx";
                wb.Save(outputPath);

                Console.WriteLine($"Cleaned workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}