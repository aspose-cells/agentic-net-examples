using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    public class BatchXlsxToPdfConverter
    {
        public static void Main()
        {
            // Folder containing the source XLSX files
            string inputFolder = "input";

            // Folder where the converted PDF files will be saved
            string outputFolder = "output";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all .xlsx files in the input folder
            string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string sourcePath in xlsxFiles)
            {
                try
                {
                    // Determine the PDF file name based on the source file name
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                    string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    // Convert the Excel file to PDF using Aspose.Cells ConversionUtility
                    ConversionUtility.Convert(sourcePath, destPath);

                    Console.WriteLine($"Converted: {sourcePath} -> {destPath}");
                }
                catch (Exception ex)
                {
                    // Log any conversion errors but continue processing other files
                    Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}