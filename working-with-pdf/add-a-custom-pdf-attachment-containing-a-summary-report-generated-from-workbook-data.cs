// Title: Embed a Generated Summary Text File as a PDF Attachment with Aspose.Cells for .NET
// Description: This example creates an Excel workbook, generates a textual summary of its data, saves the summary as a temporary file, embeds it as an OLE object, and exports the workbook to PDF with the EmbedAttachments option enabled. Temporary files are removed after the PDF is created.
// Keywords: Aspose.Cells PDF attachment | EmbedAttachments option | C# OLE object Excel | generate summary from worksheet | export workbook to PDF with attachment | Aspose.Cells PdfSaveOptions example
// Common Searches: how to embed a text file in a PDF using Aspose.Cells | Aspose.Cells PdfSaveOptions EmbedAttachments C# | add OLE object to Excel and export to PDF | create PDF with attached summary report | Aspose.Cells embed attachment example
// Developer Intent: Produce a PDF from an Excel workbook that contains a dynamically generated summary file attached via OLE embedding.
// Use Cases: Deliver a sales report PDF with an attached summary for auditors. | Provide a PDF data dictionary generated from worksheet contents. | Create an invoice PDF that includes a generated terms‑and‑conditions text file.
// AI Prompts: Show C# code to embed a generated CSV file as a PDF attachment using Aspose.Cells. | Explain the role of PdfSaveOptions.EmbedAttachments when exporting OLE objects to PDF. | Provide a version of the sample that avoids writing the summary to disk.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example creates an Excel workbook, generates a textual summary of its data, saves the summary as a temporary file, embeds it as an OLE object, and exports the workbook to PDF with the EmbedAttachments option enabled. Temporary files are removed after the PDF is created.
class AddPdfAttachmentWithSummary
{
    static void Main()
    {
        // Paths for temporary files
        string excelPath = "ReportData.xlsx";
        string summaryPath = "SummaryReport.txt";
        string pdfPath = "ReportWithSummary.pdf";

        try
        {
            // ---------- Create a workbook and add sample data ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);
            sheet.Cells["A4"].PutValue("Cherries");
            sheet.Cells["B4"].PutValue(45);

            // Save the workbook to a temporary Excel file (required for the summary generation)
            workbook.Save(excelPath);

            // ---------- Generate a summary report ----------
            string summaryText = GenerateSummary(workbook);

            // Write the summary to a text file that will be embedded as an attachment.
            File.WriteAllText(summaryPath, summaryText);

            // ---------- Embed the summary file as an OLE object ----------
            if (File.Exists(summaryPath))
            {
                try
                {
                    byte[] summaryBytes = File.ReadAllBytes(summaryPath);
                    int oleIndex = sheet.OleObjects.Add(10, 10, 200, 200, summaryBytes);
                    // Set to Unknown if specific format is not required.
                    sheet.OleObjects[oleIndex].FileFormatType = FileFormatType.Unknown;
                    sheet.OleObjects[oleIndex].DisplayAsIcon = true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to embed OLE object: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Summary file not found; skipping OLE embedding.");
            }

            // ---------- Save the workbook as PDF with embedded attachment ----------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Enable embedding of OLE attachments in the resulting PDF.
                EmbedAttachments = true
            };

            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"PDF with embedded summary created: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // ---------- Clean up temporary files ----------
            if (File.Exists(excelPath)) File.Delete(excelPath);
            if (File.Exists(summaryPath)) File.Delete(summaryPath);
        }
    }

    // Simple summary generator that creates a textual summary of the worksheet data.
    private static string GenerateSummary(Workbook workbook)
    {
        Worksheet sheet = workbook.Worksheets[0];
        int lastRow = sheet.Cells.MaxDataRow;
        int totalQuantity = 0;
        for (int row = 1; row <= lastRow; row++)
        {
            totalQuantity += sheet.Cells[row, 1].IntValue; // Column B (index 1)
        }

        return $"The report contains {lastRow} product entries with a total quantity of {totalQuantity}.";
    }
}
