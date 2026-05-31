using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    public class XlsxToPdfBatchConverter
    {
        // Converts all .xlsx files in the specified folder to PDF.
        // The ConversionUtility preserves page margins, orientation and other print settings.
        public static void ConvertFolder(string sourceFolder, string destinationFolder)
        {
            // Ensure source folder exists
            if (!Directory.Exists(sourceFolder))
                throw new DirectoryNotFoundException($"Source folder not found: {sourceFolder}");

            // Create destination folder if it does not exist
            Directory.CreateDirectory(destinationFolder);

            // Get all .xlsx files (including .xlsm, .xltx, etc. if needed)
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string excelPath in excelFiles)
            {
                // Build PDF file name with same base name
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                string pdfPath = Path.Combine(destinationFolder, fileNameWithoutExt + ".pdf");

                // Perform conversion using Aspose.Cells ConversionUtility
                // This method respects the workbook's page setup (margins, orientation, etc.)
                ConversionUtility.Convert(excelPath, pdfPath);

                Console.WriteLine($"Converted: {excelPath} -> {pdfPath}");
            }
        }

        // Example entry point
        public static void Main(string[] args)
        {
            // Example usage:
            // args[0] = source folder, args[1] = destination folder
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: XlsxToPdfBatchConverter <sourceFolder> <destinationFolder>");
                return;
            }

            try
            {
                ConvertFolder(args[0], args[1]);
                Console.WriteLine("Batch conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}