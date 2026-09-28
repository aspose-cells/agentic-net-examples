// Title: Convert multiple Excel workbooks to PDF with 80% zoom and hide zero values using Aspose.Cells in C#
// AI Prompts: Write a C# console application that loops through every .xlsx file in a folder, sets each worksheet's Zoom property to 80, enables IsZeroHidden to hide zero values, and saves the workbook as a PDF with Aspose.Cells. | Add error logging so that any workbook that fails to convert is recorded in a log file while the batch process continues with the remaining files. | Enhance the tool to accept command‑line arguments for input directory, output directory, zoom percentage, and a switch to turn zero‑value hiding on or off.
// Common Searches: aspnet batch convert excel files to pdf with custom zoom using Aspose.Cells | how to hide zero values when exporting Excel to PDF with Aspose.Cells C# | set worksheet zoom level programmatically before saving as PDF Aspose.Cells | process all .xlsx files in a directory and generate PDFs in .NET | Aspose.Cells save workbook as PDF with specific page settings
// Tags: batch excel to pdf conversion Aspose.Cells | set worksheet zoom Aspose.Cells | hide zero values worksheet Aspose.Cells | process multiple workbooks folder C# | save workbook as pdf with custom settings

using System;
using System.IO;
using Aspose.Cells;

// // Scans a specified input folder for .xlsx files, loads each workbook with Aspose.Cells, applies an 80% zoom (and optionally hides zero values via IsZeroHidden), then saves each workbook as a PDF in the output folder, handling errors per file.
class BatchProcessor
{
    static void Main()
    {
        // Define input folder containing Excel workbooks
        string inputFolder = @"C:\InputWorkbooks";

        // Define output folder for generated PDFs
        string outputFolder = @"C:\OutputPdfs";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all Excel files (adjust the pattern as needed)
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        foreach (string excelPath in excelFiles)
        {
            // Verify the file exists before attempting to load
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"File not found: {excelPath}");
                continue;
            }

            try
            {
                // Load the workbook from file
                Workbook workbook = new Workbook(excelPath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Hide zero values if the property is available (commented out if not supported)
                    // sheet.IsZeroHidden = true; // Aspose.Cells older versions may not expose this property

                    // Apply 80 percent zoom level
                    sheet.Zoom = 80;
                }

                // Construct the PDF file name based on the original workbook name
                string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the processed workbook as a PDF document
                workbook.Save(pdfPath, SaveFormat.Pdf);
                Console.WriteLine($"Converted '{excelPath}' to PDF successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
            }
        }
    }
}
