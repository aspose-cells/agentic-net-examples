using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file that represents an XLSX workbook
            string sourcePath = "workbook.json";

            // Desired output CSV file path
            string destPath = "workbook.csv";

            try
            {
                // Convert the JSON (Excel) file to CSV using default conversion options
                ConversionUtility.Convert(sourcePath, destPath);

                Console.WriteLine($"Conversion successful: '{sourcePath}' -> '{destPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}