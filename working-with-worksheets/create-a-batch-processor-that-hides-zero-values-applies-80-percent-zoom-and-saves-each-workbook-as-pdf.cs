// Title: Batch convert Excel workbooks to PDF, hide zero values and apply 80% zoom with Aspose.Cells for .NET
// AI Prompts: Generate a C# console application that scans a directory for .xlsx files, loads each workbook with Aspose.Cells, sets a custom number format ('0;-0;;@') to hide zeros, changes every worksheet's PageSetup.Zoom to 80, and saves the workbook as a PDF in a target folder. | Extend the batch converter to write a detailed CSV log that records the source file name, conversion status (success or failure), and any exception messages while continuing to process remaining files. | Add a command‑line option that skips worksheets containing only zero values from the PDF export, using Aspose.Cells APIs to detect such sheets.
// Common Searches: aspocells hide zero values in entire workbook c# | batch convert multiple xlsx files to pdf with custom zoom using Aspose.Cells | set worksheet page setup zoom to 80 percent before pdf export c# | c# suppress zero display with custom number format Aspose.Cells | process folder of Excel files and save as pdf Aspose.Cells .NET
// Tags: hide zero values Aspose.Cells | worksheet zoom percentage Aspose.Cells | batch excel to pdf conversion C# | custom number format suppress zeros .NET | process multiple workbooks Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# console program iterates over all .xlsx files in a specified input folder, loads each workbook with Aspose.Cells, applies a custom number format ('0;-0;;@') to suppress zero values workbook‑wide, sets every worksheet's PageSetup.Zoom to 80 %, and saves the result as a PDF in an output directory. It creates missing output folders, logs conversion progress, and gracefully handles missing files or runtime exceptions.
class BatchProcessor
{
    static void Main()
    {
        try
        {
            // Folder containing the source Excel workbooks
            string inputFolder = @"C:\InputWorkbooks";
            // Folder where the PDF files will be saved
            string outputFolder = @"C:\OutputPdfs";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each Excel file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Set a custom number format that suppresses zero values for the whole workbook
                    workbook.DefaultStyle.Custom = "0;-0;;@";

                    // Iterate through all worksheets in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Apply 80 percent zoom to the worksheet view (using PageSetup zoom for compatibility)
                        sheet.PageSetup.Zoom = 80;
                    }

                    // Build the output PDF file name (same as Excel file name but with .pdf extension)
                    string pdfFileName = Path.GetFileNameWithoutExtension(filePath) + ".pdf";
                    string pdfPath = Path.Combine(outputFolder, pdfFileName);

                    // Save the workbook as PDF
                    workbook.Save(pdfPath, SaveFormat.Pdf);
                    Console.WriteLine($"Converted: {Path.GetFileName(filePath)} -> {pdfFileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
