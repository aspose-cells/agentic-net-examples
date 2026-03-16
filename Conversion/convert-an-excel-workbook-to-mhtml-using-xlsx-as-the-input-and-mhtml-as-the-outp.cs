using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Desired output path for the MHTML file
            string destPath = "output.mht";

            try
            {
                // Convert the Excel workbook to MHTML using the provided ConversionUtility method
                ConversionUtility.Convert(sourcePath, destPath);

                Console.WriteLine("Conversion completed successfully.");
                Console.WriteLine($"Source: {sourcePath}");
                Console.WriteLine($"Destination: {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}