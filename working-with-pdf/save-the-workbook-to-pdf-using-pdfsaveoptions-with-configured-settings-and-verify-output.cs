// Title: Export an Aspose.Cells workbook to PDF with OnePagePerSheet and AllColumnsInOnePagePerSheet options and confirm file creation in C#
// AI Prompts: Write C# code that creates a Workbook, populates cells, configures PdfSaveOptions with OnePagePerSheet = true and AllColumnsInOnePagePerSheet = true, and saves the workbook as a PDF. | Add logic to verify the target directory exists, create it if missing, and then call workbook.Save using the configured PdfSaveOptions. | After the save operation, implement a check that uses File.Exists to confirm the PDF was generated and output its full path.
// Common Searches: Aspose.Cells C# export Excel to PDF using OnePagePerSheet option | How to keep all columns on a single PDF page per worksheet with Aspose.Cells | Create output folder before saving PDF with Aspose.Cells in .NET | Check if PDF file was generated after workbook.Save in C# | PdfSaveOptions settings for single-page-per-sheet PDF conversion Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | C# export workbook to PDF AllColumnsInOnePagePerSheet | ensure output folder exists before saving PDF | validate generated PDF file existence in .NET | single-page-per-sheet PDF conversion using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills it with sample data, configures PdfSaveOptions (OnePagePerSheet and AllColumnsInOnePagePerSheet), ensures the output directory exists, saves the workbook as a PDF, and verifies that the PDF file was successfully created.
class PdfExportExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // empty workbook

            // Add sample data to demonstrate PDF export
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export each worksheet to a separate page
                OnePagePerSheet = true,
                // Preserve the original layout
                AllColumnsInOnePagePerSheet = true
                // Note: PdfCompliance setting removed for compatibility with current library version
            };

            // Define output file path
            string outputPath = "WorkbookOutput.pdf";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            // Verify that the PDF file was created successfully
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"PDF file successfully saved to: {Path.GetFullPath(outputPath)}");
            }
            else
            {
                Console.WriteLine("Failed to create the PDF file.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
