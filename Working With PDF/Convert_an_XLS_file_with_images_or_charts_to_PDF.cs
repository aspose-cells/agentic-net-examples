using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    public class XlsToPdfConverter
    {
        /// <summary>
        /// Converts an Excel file (XLS) that may contain images or charts to a PDF document.
        /// </summary>
        /// <param name="sourcePath">Full path of the source XLS file.</param>
        /// <param name="pdfPath">Full path where the resulting PDF will be saved.</param>
        public static void ConvertXlsToPdf(string sourcePath, string pdfPath)
        {
            // Use the ConversionUtility provided by Aspose.Cells to perform the conversion.
            // This method handles all worksheet content, including images and charts.
            ConversionUtility.Convert(sourcePath, pdfPath);
        }

        // Example entry point
        public static void Main()
        {
            // Define source XLS file and destination PDF file paths
            string sourceFile = "sample_with_images_and_charts.xls";
            string destinationPdf = "output.pdf";

            try
            {
                ConvertXlsToPdf(sourceFile, destinationPdf);
                Console.WriteLine($"Conversion completed successfully. PDF saved to: {destinationPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}