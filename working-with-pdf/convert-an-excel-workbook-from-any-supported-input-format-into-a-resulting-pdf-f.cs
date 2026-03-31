using System;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    public class ExcelToPdfConverter
    {
        /// <summary>
        /// Converts an Excel file (any supported format) to PDF.
        /// </summary>
        /// <param name="sourcePath">Full path of the source Excel file.</param>
        /// <param name="pdfPath">Full path where the PDF will be saved.</param>
        public static void ConvertToPdf(string sourcePath, string pdfPath)
        {
            // Use Aspose.Cells.Utility.ConversionUtility to perform the conversion.
            // This follows the provided rule: ConversionUtility.Convert(string, string)
            ConversionUtility.Convert(sourcePath, pdfPath);
        }

        // Example usage
        public static void Main()
        {
            // Define input and output file paths.
            string sourceFile = "input.xlsx";   // can be .xls, .csv, .ods, etc.
            string outputPdf = "output.pdf";

            // Perform conversion.
            ConvertToPdf(sourceFile, outputPdf);

            Console.WriteLine($"Conversion completed: '{sourceFile}' -> '{outputPdf}'");
        }
    }
}