using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPdfAttachmentDemo
{
    class Program
    {
        static void Main()
        {
            // Create temporary files that will be embedded as attachments
            string[] tempFiles = new string[]
            {
                Path.Combine(Path.GetTempPath(), "SampleDoc.docx"),
                Path.Combine(Path.GetTempPath(), "SamplePdf.pdf"),
                Path.Combine(Path.GetTempPath(), "SampleXls.xlsx")
            };

            // Write simple content to each file
            File.WriteAllText(tempFiles[0], "This is a sample DOCX document.");
            File.WriteAllBytes(tempFiles[1], new byte[] { 0x25, 0x50, 0x44, 0x46 }); // minimal PDF header "%PDF"
            File.WriteAllText(tempFiles[2], "A1,Sample,XLSX\n"); // placeholder content for XLSX (will be ignored)

            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("PDF with Multiple Embedded Attachments");

            // Add each file as an OLE object (displayed as an icon)
            int startRow = 5;
            int startColumn = 0;
            int iconWidth = 100;
            int iconHeight = 100;

            for (int i = 0; i < tempFiles.Length; i++)
            {
                string filePath = tempFiles[i];
                byte[] fileData = File.ReadAllBytes(filePath);

                // Add OLE object at a different position for each file
                int oleIndex = sheet.OleObjects.Add(startRow + i * 5, startColumn, iconWidth, iconHeight, fileData);
                OleObject ole = sheet.OleObjects[oleIndex];

                // Set the appropriate file format type based on extension
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                switch (ext)
                {
                    case ".docx":
                        ole.FileFormatType = FileFormatType.Docx;
                        break;
                    case ".pdf":
                        ole.FileFormatType = FileFormatType.Pdf;
                        break;
                    case ".xlsx":
                        ole.FileFormatType = FileFormatType.Xlsx;
                        break;
                    default:
                        ole.FileFormatType = FileFormatType.Unknown;
                        break;
                }

                // Display as an icon with a label
                ole.DisplayAsIcon = true;
                ole.Label = Path.GetFileName(filePath);
            }

            // Configure PDF save options to embed attachments
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true
            };

            // Save the workbook as PDF with embedded attachments
            string outputPdf = Path.Combine(Path.GetTempPath(), "WorkbookWithAttachments.pdf");
            workbook.Save(outputPdf, pdfOptions);

            Console.WriteLine($"PDF saved to: {outputPdf}");
            Console.WriteLine("Embedded attachments:");
            foreach (string f in tempFiles)
                Console.WriteLine($"- {f}");

            // Clean up temporary files (optional)
            foreach (string f in tempFiles)
                File.Delete(f);
        }
    }
}