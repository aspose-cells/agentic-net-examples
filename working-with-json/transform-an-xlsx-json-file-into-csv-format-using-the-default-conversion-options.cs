using System;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file (the file may contain JSON data inside cells)
            string sourcePath = "input.xlsx";

            // Desired output CSV file path
            string destinationPath = "output.csv";

            // Convert the XLSX file to CSV using default conversion options
            // The ConversionUtility handles loading the source workbook and saving it in the target format.
            ConversionUtility.Convert(sourcePath, destinationPath);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destinationPath}'");
        }
    }
}