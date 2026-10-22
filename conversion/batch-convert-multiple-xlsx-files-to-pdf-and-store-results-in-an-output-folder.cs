// Title: Batch convert a folder of XLSX workbooks to PDF files with Aspose.Cells in C#
// Common Searches: asp.net batch convert xlsx files to pdf using Aspose.Cells | c# code to convert all Excel workbooks in a directory to PDF | how to save multiple .xlsx as PDF programmatically with Aspose.Cells | bulk Excel to PDF conversion example in .NET | process folder of Excel files to PDF with error handling C#
// Tags: Aspose.Cells XLSX to PDF batch conversion | C# bulk Excel workbook PDF export | SaveFormat.Pdf workbook saving .NET | directory based Excel to PDF automation | exception handling Aspose.Cells bulk conversion | create output folder for PDF results C#

using System;
using System.IO;
using Aspose.Cells;

namespace BatchConversion
{
    // The sample iterates through every .xlsx file in a specified input directory, loads each workbook with Aspose.Cells, and saves it as a PDF in a designated output folder, creating the folder if needed and logging success or error messages for each file.
    class BatchXlsxToPdfConverter
    {
        static void Main(string[] args)
        {
            // Define input and output directories.
            string inputFolder = @"C:\InputXlsx";
            string outputFolder = @"C:\OutputPdf";

            // Verify input directory exists.
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure the output directory exists.
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            string[] xlsxFiles;
            try
            {
                // Retrieve all .xlsx files from the input folder.
                xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error accessing input folder: {ex.Message}");
                return;
            }

            foreach (string xlsxPath in xlsxFiles)
            {
                try
                {
                    // Confirm the file exists before loading.
                    if (!File.Exists(xlsxPath))
                    {
                        Console.WriteLine($"File not found: {xlsxPath}");
                        continue;
                    }

                    // Load the XLSX workbook.
                    Workbook workbook = new Workbook(xlsxPath);

                    // Construct the output PDF file path.
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                    string pdfPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    // Save the workbook as PDF.
                    workbook.Save(pdfPath, SaveFormat.Pdf);

                    Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {fileNameWithoutExt}.pdf");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting file '{xlsxPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
