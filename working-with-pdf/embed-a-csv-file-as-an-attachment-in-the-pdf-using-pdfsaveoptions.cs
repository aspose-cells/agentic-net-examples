// Title: Embed CSV as PDF Attachment with Aspose.Cells PdfSaveOptions (C#)
// Description: The C# sample builds a workbook, generates a temporary CSV file, adds it to the sheet as an OLE object shown as an icon, sets PdfSaveOptions to embed attachments, and saves the result as a PDF that contains the CSV file.
// Keywords: Aspose.Cells PDF attachment C# | PdfSaveOptions EmbedAttachments | OLE object CSV Aspose.Cells | save workbook as PDF with embedded file | C# embed CSV in PDF | Aspose.Cells add attachment to PDF
// Common Searches: how to embed a csv file in a pdf using aspose.cells | pdfsaveoptions embedattachments example c# | add ole object to worksheet and include in pdf output | aspose.cells embed file as pdf attachment
// Developer Intent: Add a CSV file to a workbook as an OLE object and have it packaged inside the generated PDF.
// Use Cases: Distribute a financial report PDF that also carries the raw CSV data for auditors. | Create an invoice PDF that bundles the line‑item CSV export for downstream processing. | Provide a technical manual in PDF format while attaching a configuration CSV for offline analysis.
// AI Prompts: Show how to embed several CSV files as attachments in a single PDF with Aspose.Cells. | Explain the OLE object properties needed to display an icon and be included in the PDF output. | Give code for robust error handling when the CSV file is missing or inaccessible during PDF creation.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Saving;

// The C# sample builds a workbook, generates a temporary CSV file, adds it to the sheet as an OLE object shown as an icon, sets PdfSaveOptions to embed attachments, and saves the result as a PDF that contains the CSV file.
class EmbedCsvInPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("PDF with Embedded CSV Attachment");

            // Prepare a sample CSV file
            string csvPath = "sample.csv";
            string csvContent = "Name,Age\nJohn,30\nAlice,25";

            // Write CSV content (overwrite if it already exists)
            File.WriteAllText(csvPath, csvContent);

            // Ensure the CSV file exists before adding it as an OLE object
            if (!File.Exists(csvPath))
                throw new FileNotFoundException("CSV file was not created.", csvPath);

            // Read CSV bytes for embedding
            byte[] csvBytes = File.ReadAllBytes(csvPath);

            // Add the CSV file as an OLE object (attachment) to the worksheet
            int oleIndex = worksheet.OleObjects.Add(10, 10, 200, 200, csvBytes);
            OleObject oleObject = worksheet.OleObjects[oleIndex];
            oleObject.FileFormatType = FileFormatType.Csv;
            oleObject.DisplayAsIcon = true;

            // Configure PDF save options to embed attachments
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true
            };

            // Save the workbook as a PDF with the embedded CSV attachment
            string pdfPath = "PdfWithCsvAttachment.pdf";
            workbook.Save(pdfPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
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
