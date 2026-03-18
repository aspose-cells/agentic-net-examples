using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSM workbook
            string sourcePath = "input.xlsm";

            // Destination path with .json extension – determines JSON output format
            string destinationPath = "output.json";

            // Convert the XLSM file to JSON using default conversion settings.
            // The ConversionUtility automatically selects the output format based on the file extension.
            ConversionUtility.Convert(sourcePath, destinationPath);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}