// Title: Convert all XLSX files in a folder to separate PDF documents using Aspose.Cells for .NET (C#)
// Common Searches: C# code to batch convert Excel .xlsx files in a directory to PDF with Aspose.Cells | How to process multiple workbooks and save each as PDF using Aspose.Cells .NET | Aspose.Cells example for converting every .xlsx in a folder to individual PDF files | Automate folder‑wide XLSX to PDF conversion in C# with Aspose.Cells | Loop through Excel files and export each to PDF using Aspose.Cells SaveFormat.Pdf
// Tags: Aspose.Cells batch XLSX to PDF conversion | C# directory Excel to PDF processing | SaveFormat.Pdf workbook export | automated Excel PDF generation .NET | convert multiple workbooks programmatically

using System;
using System.IO;
using Aspose.Cells;

// The sample scans a given input folder for .xlsx files, creates an output folder if needed, loads each workbook with Aspose.Cells, and saves it as a PDF using SaveFormat.Pdf, preserving the original file name. It logs each successful conversion and any errors, enabling batch conversion of Excel workbooks to individual PDF files in C#.
class BatchXlsxToPdfConverter
{
    static void Main(string[] args)
    {
        // Input directory containing XLSX files
        string inputFolder = @"C:\InputXlsx";
        // Output directory for generated PDF files
        string outputFolder = @"C:\OutputPdf";

        // Ensure output directory exists
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        // Get all .xlsx files in the input directory (non-recursive)
        string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string xlsxPath in xlsxFiles)
        {
            try
            {
                // Load the workbook from the XLSX file
                Workbook workbook = new Workbook(xlsxPath);

                // Build the output PDF file path (same name, .pdf extension)
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                string pdfPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                // Save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {Path.GetFileName(pdfPath)}");
            }
            catch (Exception ex)
            {
                // Log any errors but continue processing other files
                Console.WriteLine($"Error converting '{Path.GetFileName(xlsxPath)}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
