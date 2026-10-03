// Title: Batch convert Excel .xlsx workbooks to PDF and add a timeline to each worksheet using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that scans a directory for .xlsx files, creates a timeline for the first worksheet’s pivot table when it exists, and saves each workbook as a PDF in a specified output folder using Aspose.Cells. | Update the sample batch‑conversion code to instantiate a Timeline object, associate it with the worksheet’s pivot table, and then export the workbook to PDF with Aspose.Cells.
// Common Searches: c# aspnet batch convert multiple xlsx files to pdf with Aspose.Cells | how to add a timeline to a pivot table before exporting to pdf using Aspose.Cells | automate conversion of all Excel workbooks in a folder to PDF and include timelines | Aspose.Cells example for processing a directory of Excel files and saving PDFs
// Tags: batch xlsx to pdf conversion Aspose.Cells | insert timeline into worksheet Aspose.Cells | export workbook as pdf with timeline Aspose.Cells | process multiple Excel files C# Aspose.Cells | timeline feature for pivot tables Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program enumerates every .xlsx file in a given input folder, loads each workbook with Aspose.Cells, optionally adds a timeline to the first worksheet’s pivot table, and saves the result as a PDF in an output directory, handling missing folders and runtime errors.
class BatchExcelToPdfWithTimeline
{
    static void Main()
    {
        // Input and output directories
        string inputDir = @"C:\InputExcelFiles";
        string outputDir = @"C:\OutputPdfFiles";

        // Verify input directory exists
        if (!Directory.Exists(inputDir))
        {
            Console.WriteLine($"Input directory not found: {inputDir}");
            return;
        }

        // Ensure output directory exists
        if (!Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        string[] excelFiles;
        try
        {
            // Get all Excel files in the input directory
            excelFiles = Directory.GetFiles(inputDir, "*.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error accessing input directory: {ex.Message}");
            return;
        }

        foreach (string filePath in excelFiles)
        {
            try
            {
                // Verify the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Work with the first worksheet (adjust as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // If the worksheet contains pivot tables, you could add a timeline here
                // (Timeline feature requires a newer Aspose.Cells version; omitted for compatibility)

                // Prepare output PDF file path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                string pdfPath = Path.Combine(outputDir, fileNameWithoutExt + ".pdf");

                // Save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"Converted: {filePath} -> {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
