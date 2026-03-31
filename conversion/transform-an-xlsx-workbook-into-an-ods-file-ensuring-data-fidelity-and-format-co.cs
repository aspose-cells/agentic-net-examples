using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Desired path for the resulting ODS file
            string destinationPath = "output.ods";

            // Convert the XLSX workbook to ODS format.
            // The ConversionUtility handles loading, conversion, and saving internally,
            // preserving data fidelity and applying default format compliance.
            ConversionUtility.Convert(sourcePath, destinationPath);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destinationPath}'");
        }
    }
}