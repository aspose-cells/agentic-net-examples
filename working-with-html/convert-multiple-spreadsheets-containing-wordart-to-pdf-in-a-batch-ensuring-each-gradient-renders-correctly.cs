using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Folder containing the source Excel files with WordArt
                string sourceFolder = @"C:\InputSpreadsheets";

                // Verify source folder exists
                if (!Directory.Exists(sourceFolder))
                {
                    Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                    return;
                }

                // Folder where the converted PDF files will be saved
                string outputFolder = @"C:\OutputPDFs";

                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Get all Excel files (you can adjust the pattern if needed)
                string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

                foreach (string excelPath in excelFiles)
                {
                    // Verify the file still exists before processing
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found: {excelPath}");
                        continue;
                    }

                    try
                    {
                        // Build the PDF file name based on the Excel file name
                        string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                        string pdfPath = Path.Combine(outputFolder, pdfFileName);

                        // Convert the Excel workbook to PDF.
                        // ConversionUtility handles rendering of all objects, including WordArt gradients.
                        ConversionUtility.Convert(excelPath, pdfPath);

                        Console.WriteLine($"Successfully converted: {excelPath} -> {pdfPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting '{excelPath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}