// Title: C# batch conversion of XLSX workbooks to PDF using Aspose.Cells
// Common Searches: C# code to convert all .xlsx files in a directory to PDF with Aspose.Cells | How to loop through Excel files and save each as PDF using Aspose.Cells .NET | Batch processing multiple Excel workbooks to PDF in a console application | Aspose.Cells example for converting a folder of XLSX to PDF in C# | Automate Excel to PDF conversion for many files using Aspose.Cells library
// Tags: Aspose.Cells batch XLSX to PDF conversion | C# iterate Excel files and export to PDF | Workbook.Save with SaveFormat.Pdf in .NET | process multiple Excel workbooks in a folder | automated Excel to PDF conversion .NET

using System;
using System.IO;
using Aspose.Cells;

namespace BatchXlsxToPdf
{
    // The console program scans a specified input folder for .xlsx files, loads each workbook with Aspose.Cells, and saves it as a PDF with the same base name in an output folder, handling missing files and errors while logging progress.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output directories (adjust paths as needed)
            string inputFolder = @"C:\InputXlsx";
            string outputFolder = @"C:\OutputPdf";

            // Verify input directory exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                Console.WriteLine("Please create the folder and add .xlsx files before running the program.");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            string[] xlsxFiles;
            try
            {
                // Get all .xlsx files in the input folder
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
                    // Verify the file still exists before loading
                    if (!File.Exists(xlsxPath))
                    {
                        Console.WriteLine($"File not found (skipped): {xlsxPath}");
                        continue;
                    }

                    // Load the workbook from the XLSX file
                    Workbook workbook = new Workbook(xlsxPath);

                    // Build the output PDF file path (same name, .pdf extension)
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                    string pdfPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    // Save the workbook as PDF
                    workbook.Save(pdfPath, SaveFormat.Pdf);

                    Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {fileNameWithoutExt}.pdf");
                }
                catch (Exception ex)
                {
                    // Log any errors but continue processing other files
                    Console.WriteLine($"Error converting {Path.GetFileName(xlsxPath)}: {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
