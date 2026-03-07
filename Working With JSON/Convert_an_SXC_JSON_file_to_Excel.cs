using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source SXC file (StarOffice Calc Spreadsheet)
            string sourcePath = "input.sxc";

            // Desired output Excel file path (XLSX format)
            string outputPath = "output.xlsx";

            // Ensure the source file exists; if not, create a simple workbook and save as SXC
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Name = "Sheet1";
                tempWb.Save(sourcePath, SaveFormat.SXC);
            }

            // Convert the SXC file to Excel using Aspose.Cells ConversionUtility
            ConversionUtility.Convert(sourcePath, outputPath);

            // Load the newly created workbook to verify the conversion succeeded
            var workbook = new Workbook(outputPath);
            Console.WriteLine($"Conversion completed. Workbook contains {workbook.Worksheets.Count} worksheet(s).");
        }
    }
}