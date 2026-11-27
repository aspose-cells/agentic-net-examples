// Title: Generate a ZIP archive with a cover PDF and individual worksheet PDFs from an Excel workbook using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells export each worksheet to a separate PDF file | How to create a cover PDF that shows workbook details with Aspose.Cells | Package multiple PDF files into a ZIP archive in .NET | Use PdfSaveOptions SheetSet to save a single sheet as PDF in C# | Create temporary PDF attachments from Excel sheets and zip them
// Tags: Aspose.Cells export worksheet to PDF C# | PdfSaveOptions SheetSet single sheet export | generate cover PDF with workbook summary Aspose.Cells | zip multiple PDFs .NET C# | memory stream PDF creation Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace EmbedWorksheetsAsPdfAttachments
{
    // The example loads an Excel workbook, iterates through its worksheets, and uses PdfSaveOptions with a SheetSet to render each sheet as an individual PDF stored in memory. It then adds workbook information to the first sheet to create a cover PDF, writes all PDFs to temporary files, and bundles the cover PDF plus the worksheet PDFs into a ZIP archive, cleaning up temporary files afterwards.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Input Excel file – change the path as needed
                const string inputExcelPath = "input.xlsx";

                if (!File.Exists(inputExcelPath))
                {
                    Console.WriteLine($"Error: The file '{inputExcelPath}' does not exist.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputExcelPath);
                int sheetCount = workbook.Worksheets.Count;

                // Prepare a list that will hold each worksheet's PDF representation
                List<(string FileName, byte[] Content)> worksheetPdfAttachments = new List<(string, byte[])>();

                // Convert each worksheet to an individual PDF stored in memory
                for (int i = 0; i < sheetCount; i++)
                {
                    // Configure PDF save options to render only the current sheet
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        // Use SheetSet to limit rendering to a single sheet by index
                        SheetSet = new SheetSet(new int[] { i })
                    };

                    using (MemoryStream ms = new MemoryStream())
                    {
                        // Save the selected sheet to the memory stream
                        workbook.Save(ms, pdfOptions);
                        ms.Position = 0;

                        // Store the PDF bytes together with a meaningful file name
                        string attachmentName = $"{workbook.Worksheets[i].Name}.pdf";
                        worksheetPdfAttachments.Add((attachmentName, ms.ToArray()));
                    }

                    Console.WriteLine($"Worksheet '{workbook.Worksheets[i].Name}' saved as PDF attachment.");
                }

                // -----------------------------------------------------------------
                // Create a simple cover PDF (first worksheet with explanatory text)
                // -----------------------------------------------------------------
                // We'll reuse the first worksheet to generate a cover page.
                string coverFileName = "Cover.pdf";
                using (MemoryStream coverStream = new MemoryStream())
                {
                    // Add explanatory text to the first worksheet
                    Worksheet coverSheet = workbook.Worksheets[0];
                    coverSheet.Cells["A1"].PutValue($"Workbook: {Path.GetFileName(inputExcelPath)}");
                    coverSheet.Cells["A2"].PutValue($"Total worksheets: {sheetCount}");
                    coverSheet.Cells["A4"].PutValue("Each worksheet is attached as a separate PDF for detailed review.");

                    // Save as PDF
                    workbook.Save(coverStream, new PdfSaveOptions());
                    File.WriteAllBytes(coverFileName, coverStream.ToArray());
                }

                // -----------------------------------------------------------------
                // Package cover PDF and worksheet PDFs into a ZIP archive
                // -----------------------------------------------------------------
                const string outputZipPath = "Workbook_With_Worksheet_Attachments.zip";

                // Ensure any existing archive is removed
                if (File.Exists(outputZipPath))
                {
                    File.Delete(outputZipPath);
                }

                using (FileStream zipToOpen = new FileStream(outputZipPath, FileMode.Create))
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
                {
                    // Add cover PDF
                    archive.CreateEntryFromFile(coverFileName, coverFileName);

                    // Add each worksheet PDF
                    foreach (var attachment in worksheetPdfAttachments)
                    {
                        string tempPath = Path.Combine(Path.GetTempPath(), attachment.FileName);
                        File.WriteAllBytes(tempPath, attachment.Content);
                        archive.CreateEntryFromFile(tempPath, attachment.FileName);
                        File.Delete(tempPath);
                    }
                }

                // Clean up temporary cover PDF
                if (File.Exists(coverFileName))
                {
                    File.Delete(coverFileName);
                }

                Console.WriteLine($"ZIP archive with embedded worksheet PDFs saved to '{outputZipPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
