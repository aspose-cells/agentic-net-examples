using System;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class XlsmJsonToCsvConverter
    {
        /// <summary>
        /// Converts an XLSM file (or a JSON representation of an XLSM workbook) to CSV
        /// using Aspose.Cells' default conversion options.
        /// </summary>
        /// <param name="sourcePath">Full path of the source XLSM/JSON file.</param>
        /// <param name="csvPath">Full path where the resulting CSV file will be saved.</param>
        public static void Run(string sourcePath, string csvPath)
        {
            try
            {
                // Perform conversion using the default options.
                // The ConversionUtility handles loading the source workbook and saving it as CSV.
                ConversionUtility.Convert(sourcePath, csvPath);

                Console.WriteLine($"Conversion completed successfully: '{sourcePath}' → '{csvPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }

        // Example usage
        public static void Main()
        {
            // Replace these paths with actual file locations.
            string sourceFile = "input.xlsm.json";   // Source XLSM JSON file
            string outputCsv = "output.csv";         // Desired CSV output file

            Run(sourceFile, outputCsv);
        }
    }
}