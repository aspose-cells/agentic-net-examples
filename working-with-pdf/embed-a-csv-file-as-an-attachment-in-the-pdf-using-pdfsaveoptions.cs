using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Needed for OleObject
using Aspose.Cells.Rendering;

class EmbedCsvInPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add a title
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("PDF with Embedded CSV Example");

            // Prepare a sample CSV file to embed
            string csvPath = "sample.csv";
            try
            {
                File.WriteAllText(csvPath, "Name,Score\nAlice,85\nBob,92");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create CSV file: {ex.Message}");
                return;
            }

            // Ensure the CSV file exists before embedding it as an OLE object
            if (File.Exists(csvPath))
            {
                try
                {
                    // Add the CSV as an OLE object (attachment) to the worksheet
                    int oleIndex = sheet.OleObjects.Add(10, 10, 200, 200, File.ReadAllBytes(csvPath));
                    OleObject ole = sheet.OleObjects[oleIndex];
                    ole.FileFormatType = FileFormatType.Csv;   // specify CSV format
                    ole.DisplayAsIcon = true;                 // show as an icon
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add OLE object: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("CSV file not found; skipping OLE embedding.");
            }

            // Configure PDF save options to embed OLE attachments
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true
            };

            // Save the workbook as PDF with the embedded CSV attachment
            string pdfPath = "WorkbookWithCsvAttachment.pdf";
            try
            {
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{pdfPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save PDF: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
        finally
        {
            // Clean up the temporary CSV file if it exists
            string csvPath = "sample.csv";
            if (File.Exists(csvPath))
            {
                try
                {
                    File.Delete(csvPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to delete temporary CSV file: {ex.Message}");
                }
            }
        }
    }
}