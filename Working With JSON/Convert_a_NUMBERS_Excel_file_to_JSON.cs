using System;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Apple Numbers file
            string sourcePath = "sample.numbers";

            // Desired output JSON file path
            string jsonPath = "sample.json";

            try
            {
                // Convert the Numbers file directly to JSON using the utility method
                ConversionUtility.Convert(sourcePath, jsonPath);

                Console.WriteLine($"Conversion succeeded: '{sourcePath}' → '{jsonPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}