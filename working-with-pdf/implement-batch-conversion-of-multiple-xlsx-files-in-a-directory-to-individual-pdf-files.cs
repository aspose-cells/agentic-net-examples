using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchXlsxToPdf
{
    // Author: Aspose.Cells .NET example – batch conversion of XLSX files to PDF
    class Program
    {
        static void Main()
        {
            // Folder containing source XLSX files
            string sourceFolder = @"C:\InputXlsx";

            // Folder where PDF files will be saved
            string outputFolder = @"C:\OutputPdf";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all .xlsx files in the source folder
            string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx");

            foreach (string xlsxPath in xlsxFiles)
            {
                // Build the corresponding PDF file name
                string pdfFileName = Path.GetFileNameWithoutExtension(xlsxPath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Load options specific to XLSX files
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Save options for PDF output (example: fit all columns on one page per sheet)
                PdfSaveOptions saveOptions = new PdfSaveOptions
                {
                    AllColumnsInOnePagePerSheet = true
                };

                // Convert the Excel file to PDF using the utility method
                ConversionUtility.Convert(xlsxPath, loadOptions, pdfPath, saveOptions);

                Console.WriteLine($"Converted '{Path.GetFileName(xlsxPath)}' to PDF.");
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}