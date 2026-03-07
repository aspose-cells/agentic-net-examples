using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file that represents an XLSM workbook
            string sourceJsonPath = "input.json";

            // Desired output Excel file (XLSM format to preserve macros if any)
            string outputExcelPath = "output.xlsm";

            // Use Aspose.Cells ConversionUtility to convert the JSON file to an Excel workbook.
            // The method automatically determines the format from the file extensions.
            ConversionUtility.Convert(sourceJsonPath, outputExcelPath);

            Console.WriteLine($"Conversion completed: '{sourceJsonPath}' -> '{outputExcelPath}'");
        }
    }
}