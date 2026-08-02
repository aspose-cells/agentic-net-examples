using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

class Program
{
    static void Main()
    {
        // Create a new workbook and add worksheets
        Workbook workbook = new Workbook();
        Worksheet summary = workbook.Worksheets[0];
        summary.Name = "Summary";
        summary.Cells["A1"].PutValue("This is the summary sheet.");

        Worksheet details = workbook.Worksheets.Add("Details");
        details.Cells["A1"].PutValue("Detailed data goes here.");

        Worksheet analysis = workbook.Worksheets.Add("Analysis");
        analysis.Cells["A1"].PutValue("Analysis results.");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // This property embeds OLE object attachments present in the workbook.
            // Directly embedding each worksheet as a separate PDF attachment is not supported
            // by the current Aspose.Cells API; the whole workbook can be embedded via this flag.
            EmbedAttachments = true
        };

        // Save the workbook as a PDF file
        workbook.Save("MultiSheetAttachments.pdf", pdfOptions);
    }
}

// Author note: Aspose.Cells currently provides EmbedAttachments to embed OLE objects.
// Embedding each worksheet as an individual PDF attachment is not available in the documented API.