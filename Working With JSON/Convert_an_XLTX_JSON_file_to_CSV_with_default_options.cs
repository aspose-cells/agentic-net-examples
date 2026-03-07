using System;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLTX file (Excel OpenXML template)
            string sourcePath = "sample.xltx";

            // Desired output CSV file path
            string destPath = "sample.csv";

            try
            {
                // Convert the XLTX file to CSV using default options.
                // The ConversionUtility will infer the formats from the file extensions.
                ConversionUtility.Convert(sourcePath, destPath);

                Console.WriteLine($"Conversion successful: \"{sourcePath}\" -> \"{destPath}\"");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}