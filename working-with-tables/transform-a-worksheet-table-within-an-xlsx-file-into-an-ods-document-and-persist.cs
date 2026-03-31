using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file containing the worksheet table
            string sourcePath = "input.xlsx";

            // Desired path for the resulting ODS file
            string destinationPath = "output.ods";

            // Convert the Excel file to ODS format using Aspose.Cells ConversionUtility
            // This utilizes the provided Convert(string, string) method as required by the rules
            ConversionUtility.Convert(sourcePath, destinationPath);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destinationPath}'");
        }
    }
}