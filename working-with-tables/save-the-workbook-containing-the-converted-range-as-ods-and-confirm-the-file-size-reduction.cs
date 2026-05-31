using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsOdsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and fill it with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a large range to make file size noticeable
            for (int row = 0; row < 2000; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define file names for the original XLSX and the ODS output
            string xlsxPath = "SampleData.xlsx";
            string odsPath = "SampleData.ods";

            // Save the workbook as XLSX (default options)
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            // Create ODS save options (using default constructor)
            OdsSaveOptions odsOptions = new OdsSaveOptions
            {
                // Example: set generator type to LibreOffice (optional)
                GeneratorType = OdsGeneratorType.LibreOffice
            };

            // Save the same workbook as ODS using the options
            workbook.Save(odsPath, odsOptions);

            // Get file sizes
            long xlsxSize = new FileInfo(xlsxPath).Length;
            long odsSize = new FileInfo(odsPath).Length;

            // Output sizes and reduction percentage
            Console.WriteLine($"XLSX size: {xlsxSize} bytes");
            Console.WriteLine($"ODS size:  {odsSize} bytes");

            if (xlsxSize > 0)
            {
                double reduction = (double)(xlsxSize - odsSize) / xlsxSize * 100;
                Console.WriteLine($"File size reduction: {reduction:F2}%");
            }
            else
            {
                Console.WriteLine("Original XLSX file size is zero; cannot compute reduction.");
            }

            // Clean up
            workbook.Dispose();
        }
    }
}