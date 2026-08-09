// Title: C# – Convert an Aspose.Cells Workbook to PDF with Author Metadata
// Description: Demonstrates how to set the workbook author using Workbook.Settings and BuiltInDocumentProperties, configure PdfSaveOptions to export standard properties, and save the workbook as a PDF that embeds the author information.
// Keywords: Aspose.Cells PDF conversion C# | set author property Aspose.Cells | Workbook.Settings.Author | PdfSaveOptions custom properties | embed author metadata PDF | Aspose.Cells document properties | C# Excel to PDF with metadata | export built‑in properties Aspose.Cells | PDF metadata author Aspose | Aspose.Cells .NET
// Common Searches: How to set author metadata when saving a workbook as PDF with Aspose.Cells | Aspose.Cells C# export built‑in document properties to PDF | Convert Excel to PDF and keep author information using Aspose.Cells | PdfSaveOptions to include author property in PDF | C# code to add creator name to PDF generated from Excel
// Developer Intent: Add author metadata to a workbook and generate a PDF that retains it.
// Use Cases: Create audit‑ready PDF reports that automatically display the creator’s name. | Batch‑process Excel templates into PDFs while preserving author and timestamp metadata for compliance. | Automate document generation where each PDF must carry the content creator’s identity for traceability.
// AI Prompts: Show C# code that sets the workbook author and exports it to PDF using Aspose.Cells. | Explain how to configure PdfSaveOptions to include built‑in properties such as Author in the resulting PDF. | Provide a step‑by‑step guide to verify that the author metadata appears in the PDF after conversion.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to set the workbook author using Workbook.Settings and BuiltInDocumentProperties, configure PdfSaveOptions to export standard properties, and save the workbook as a PDF that embeds the author information.
class ConvertWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the author property (content creator) using WorkbookSettings
        workbook.Settings.Author = "John Doe";

        // Also set the built‑in document property for completeness
        workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";

        // Add some sample data to demonstrate the workbook
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Workbook");
        sheet.Cells["A2"].PutValue("Created by: " + workbook.Settings.Author);
        sheet.Cells["A3"].PutValue(DateTime.Now.ToString());

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export built‑in custom properties (including Author) to the PDF file
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // Save the workbook as a PDF document
        workbook.Save("SampleOutput.pdf", pdfOptions);
    }
}
