// Title: Add Background Image to All Worksheets and Export Each Sheet as PDF with Aspose.Cells (C#)
// Description: A C# example that loads an Excel file, sets a JPEG as the background image for every worksheet, creates a temporary workbook for each sheet, and converts each sheet to a separate PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells background image C# | export worksheet to PDF Aspose.Cells | add watermark to Excel sheets | convert each Excel sheet to PDF | temporary workbook PDF conversion | C# Excel to PDF with background | Aspose.Cells SaveOptions PDF
// Common Searches: set background picture for all worksheets Aspose.Cells | export each Excel sheet as individual PDF C# | Aspose.Cells add watermark and convert to PDF | C# code to apply background image to Excel workbook | how to use ConversionUtility for PDF export
// Developer Intent: Load an Excel workbook, apply a single background image to all worksheets, and generate a separate PDF file for each sheet.
// Use Cases: Brand every worksheet with a company logo before sharing PDFs with clients. | Create department‑specific PDF reports while preserving a watermark background. | Automate batch processing to add a background image and archive each sheet as an individual PDF.
// AI Prompts: Write C# code that adds a background image to all worksheets and saves each sheet directly to PDF without intermediate files using Aspose.Cells. | Show how to use PdfSaveOptions with Aspose.Cells to export worksheets to PDF while keeping the background image. | Explain performance‑friendly techniques for handling large workbooks when adding background images and converting each sheet to PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsBackgroundPdfExport
{
    // A C# example that loads an Excel file, sets a JPEG as the background image for every worksheet, creates a temporary workbook for each sheet, and converts each sheet to a separate PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths for the source Excel file and the background image
                string excelPath = "input.xlsx";
                string imagePath = "background.jpg";

                // Verify that the required files exist
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                    return;
                }

                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Error: Image file not found at '{imagePath}'.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Read the background image into a byte array
                byte[] backgroundData = File.ReadAllBytes(imagePath);

                // Apply the background image to every worksheet
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    ws.BackgroundImage = backgroundData;
                }

                // Directories for temporary files and final PDFs
                string tempDir = "temp";
                string outputDir = "output";
                Directory.CreateDirectory(tempDir);
                Directory.CreateDirectory(outputDir);

                // Export each worksheet to a separate PDF file
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    // Create a temporary workbook containing only the current sheet
                    Workbook tempWb = new Workbook();

                    // Ensure the temporary workbook has at least one worksheet
                    Worksheet targetSheet = tempWb.Worksheets[0];

                    // Copy the current worksheet into the temporary workbook
                    workbook.Worksheets[i].Copy(targetSheet);

                    // If more than one sheet exists (unlikely), remove extras
                    while (tempWb.Worksheets.Count > 1)
                    {
                        tempWb.Worksheets.RemoveAt(1);
                    }

                    // Save the temporary workbook to an intermediate Excel file
                    string tempExcelPath = Path.Combine(tempDir, $"Sheet_{i + 1}.xlsx");
                    tempWb.Save(tempExcelPath);

                    // Convert the intermediate Excel file to PDF
                    string pdfPath = Path.Combine(outputDir, $"Sheet_{i + 1}.pdf");
                    ConversionUtility.Convert(tempExcelPath, pdfPath);

                    // Clean up the intermediate file
                    if (File.Exists(tempExcelPath))
                    {
                        File.Delete(tempExcelPath);
                    }
                }

                Console.WriteLine("All sheets have been exported to PDFs with background images.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
