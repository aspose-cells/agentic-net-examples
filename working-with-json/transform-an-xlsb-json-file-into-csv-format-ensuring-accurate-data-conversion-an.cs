using System;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSB file (binary Excel workbook)
            string sourcePath = "input.xlsb";

            // Desired output CSV file path
            string destinationPath = "output.csv";

            try
            {
                // Convert the XLSB workbook to CSV format using Aspose.Cells ConversionUtility
                ConversionUtility.Convert(sourcePath, destinationPath);

                Console.WriteLine($"Conversion successful: '{sourcePath}' -> '{destinationPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}