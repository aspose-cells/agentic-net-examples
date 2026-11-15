// Title: C# batch conversion of Excel workbooks to PDF with 80% worksheet zoom using Aspose.Cells
// Common Searches: how to convert multiple .xlsx files to PDF with Aspose.Cells in C# | set worksheet zoom level before saving workbook as PDF using Aspose.Cells .NET | batch process Excel files in a folder and export to PDF with custom zoom C# | Aspose.Cells example for applying 80% zoom to all worksheets programmatically | C# script to convert all Excel workbooks in a directory to PDF using Aspose.Cells
// Tags: batch Excel to PDF conversion Aspose.Cells .NET | apply worksheet zoom Aspose.Cells C# | export workbook as PDF with custom zoom | process multiple Excel files programmatically | save Excel workbook to PDF using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace BatchProcessor
{
    // The C# program iterates through all .xlsx files in a specified input folder, loads each workbook with Aspose.Cells, sets each worksheet's Zoom property to 80%, and saves the workbook as a PDF file in an output folder.
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing Excel workbooks
            string inputFolder = @"C:\InputWorkbooks";
            // Output folder for generated PDFs
            string outputFolder = @"C:\OutputPdfs";

            // Ensure input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Process each Excel file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                try
                {
                    // Verify the file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found, skipping: {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Apply 80% zoom to every worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        sheet.Zoom = 80;
                    }

                    // Build output PDF file name
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                    string pdfPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    // Save the workbook as PDF
                    workbook.Save(pdfPath, SaveFormat.Pdf);

                    Console.WriteLine($"Converted '{filePath}' to PDF successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
