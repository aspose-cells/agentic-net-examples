using System;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonToExcel
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file (XLTX JSON representation)
            string sourceJsonPath = "input.json";

            // Desired output Excel file path (XLSX format)
            string outputExcelPath = "output.xlsx";

            // Convert the JSON file to an Excel workbook using Aspose.Cells ConversionUtility
            // This method automatically detects the source format and creates the appropriate Excel file.
            ConversionUtility.Convert(sourceJsonPath, outputExcelPath);

            Console.WriteLine($"Conversion completed: '{sourceJsonPath}' -> '{outputExcelPath}'");
        }
    }
}