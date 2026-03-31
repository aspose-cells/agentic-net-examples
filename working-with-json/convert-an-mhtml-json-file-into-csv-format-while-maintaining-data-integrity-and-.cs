using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace MhtmlToCsvConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source MHTML file that contains JSON data
            string sourcePath = "input.mhtml";

            // Desired output CSV file path
            string destinationPath = "output.csv";

            try
            {
                // Convert the MHTML file directly to CSV.
                // Aspose.Cells automatically detects the source format (MHTML) and saves as CSV.
                ConversionUtility.Convert(sourcePath, destinationPath);

                Console.WriteLine($"Conversion successful: '{sourcePath}' → '{destinationPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}