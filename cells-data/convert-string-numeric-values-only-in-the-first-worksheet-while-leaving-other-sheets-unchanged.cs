using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConvertStringToNumericInFirstSheet
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (index 0)
                Worksheet firstSheet = workbook.Worksheets[0];

                // Convert string values that can be interpreted as numbers to numeric values
                firstSheet.Cells.ConvertStringToNumericValue();

                // Save the modified workbook to the output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display the error message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}