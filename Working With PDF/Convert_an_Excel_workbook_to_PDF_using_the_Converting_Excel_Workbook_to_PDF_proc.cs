using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Desired output PDF file path
            string destPath = "output.pdf";

            try
            {
                // Convert the Excel workbook to PDF using the built‑in ConversionUtility
                ConversionUtility.Convert(sourcePath, destPath);

                Console.WriteLine("Conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}