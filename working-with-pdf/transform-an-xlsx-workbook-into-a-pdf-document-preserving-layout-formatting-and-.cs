using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    public class XlsxToPdfConverter
    {
        /// <summary>
        /// Converts an existing XLSX workbook to PDF while preserving layout, formatting, and embedded content.
        /// </summary>
        /// <param name="sourcePath">Full path to the source XLSX file.</param>
        /// <param name="destPath">Full path where the resulting PDF will be saved.</param>
        public static void ConvertXlsxToPdf(string sourcePath, string destPath)
        {
            // Use the provided ConversionUtility.Convert method (rule) to perform the conversion.
            // This method handles loading the workbook, preserving all visual aspects,
            // and saving it directly as a PDF file.
            ConversionUtility.Convert(sourcePath, destPath);
        }

        // Example usage
        public static void Main()
        {
            // Define source XLSX and target PDF file paths
            string sourceFile = "input.xlsx";
            string outputFile = "output.pdf";

            // Perform conversion
            ConvertXlsxToPdf(sourceFile, outputFile);

            Console.WriteLine($"Conversion completed: '{sourceFile}' -> '{outputFile}'");
        }
    }
}