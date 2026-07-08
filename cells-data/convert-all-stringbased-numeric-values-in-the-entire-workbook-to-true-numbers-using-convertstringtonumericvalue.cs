using System;
using Aspose.Cells;

namespace ConvertStringNumbers
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (replace with your actual file)
            string inputPath = "input.xlsx";

            // Path where the converted workbook will be saved
            string outputPath = "output_converted.xlsx";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Convert all string values that can be interpreted as numbers
                // to true numeric values within the current worksheet.
                sheet.Cells.ConvertStringToNumericValue();
            }

            // Save the modified workbook
            workbook.Save(outputPath);

            Console.WriteLine($"Conversion complete. Saved to '{outputPath}'.");
        }
    }
}