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

            // Desired output PDF file path
            string outputPath = "output.pdf";

            try
            {
                // Convert the Excel file to PDF using the provided ConversionUtility method
                // This method internally loads the workbook and saves it as PDF.
                ConversionUtility.Convert(sourcePath, outputPath);

                Console.WriteLine($"Conversion completed successfully. PDF saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}