using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRangeAddressDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a range representing cells A1:C5 (use fully qualified type to avoid ambiguity)
                Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:C5");

                // Retrieve and display the address of the created range
                string rangeAddress = range.Address;
                Console.WriteLine("The address of the range is: " + rangeAddress);

                // Ensure the output directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (if any changes were made)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}