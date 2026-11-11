// Title: Batch convert Excel workbooks (xls, xlsx, xlsm) to PDF in C# using Aspose.Cells
// Common Searches: C# Aspose.Cells batch convert all Excel files in a folder to PDF | How to loop through .xlsx files and save each as PDF with Aspose.Cells | Automate directory‑wide Excel to PDF conversion in .NET using Aspose.Cells | Save each workbook as PDF while preserving original file name in C# | Aspose.Cells convert multiple workbooks to PDF without opening Excel
// Tags: Aspose.Cells batch workbook to PDF conversion | C# iterate Excel files export PDF | SaveFormat.Pdf usage with Aspose.Cells | process multiple .xls .xlsx .xlsm files .NET | directory traversal Excel to PDF conversion

using System;
using System.IO;
using Aspose.Cells;

// The example scans a specified input directory for Excel files (xls, xlsx, xlsm), loads each workbook with Aspose.Cells, converts it to PDF using the SaveFormat.Pdf option, and writes the PDF to an output folder with the same base name, handling missing files and runtime errors.
class BatchExcelToPdf
{
    static void Main()
    {
        // Folder containing the Excel files to process
        string inputFolder = @"C:\InputExcelFiles";
        // Folder where the resulting PDF files will be saved
        string outputFolder = @"C:\OutputPdfFiles";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files (xls, xlsx, xlsm) in the input folder
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            // Filter supported Excel extensions
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                continue;

            // Verify the file still exists before processing
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook (default options)
                Workbook workbook = new Workbook(filePath);

                // Build the output PDF file path (same name, .pdf extension)
                string pdfFileName = Path.GetFileNameWithoutExtension(filePath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to PDF.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
