// Title: Create a PDF report with worksheet hyperlinks using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to build a two‑sheet Excel workbook, add hyperlinks that point to the same workbook and specific cells, save the workbook as an XLSX file, and export it to PDF while preserving the links by using PdfSaveOptions.ExportDocumentStructure.
// Keywords: Aspose.Cells PDF hyperlink | C# export Excel to PDF with links | PdfSaveOptions ExportDocumentStructure | embed worksheet hyperlink in PDF | .NET Aspose.Cells example | convert Excel to PDF preserving hyperlinks | hyperlink to sheet in PDF
// Common Searches: Aspose.Cells keep Excel hyperlinks when converting to PDF | C# add hyperlink to worksheet and export to PDF | PdfSaveOptions retain links in PDF | how to embed sheet link in PDF using Aspose.Cells | export multi‑sheet workbook to PDF with active links
// Developer Intent: Generate a PDF document from an Excel workbook where each sheet contains a clickable link that opens the corresponding worksheet in the original Excel file.
// Use Cases: Add a sheet‑specific hyperlink to an Excel workbook and produce a PDF that lets users jump back to the source worksheet. | Create summary and detail worksheets, save them as XLSX, then deliver a PDF version with active navigation links. | Leverage PdfSaveOptions.ExportDocumentStructure to maintain document hierarchy and hyperlink functionality in PDFs generated from Excel files.
// AI Prompts: Write C# code with Aspose.Cells that adds a hyperlink to each worksheet pointing to the same workbook and cell, then saves the workbook as PDF preserving the links. | Explain how PdfSaveOptions.ExportDocumentStructure affects hyperlink behavior in PDFs created from Excel workbooks. | Provide step‑by‑step instructions to build a summary and details sheet, embed sheet‑specific hyperlinks, and export the workbook to a PDF with clickable links using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to build a two‑sheet Excel workbook, add hyperlinks that point to the same workbook and specific cells, save the workbook as an XLSX file, and export it to PDF while preserving the links by using PdfSaveOptions.ExportDocumentStructure.
class PdfReportWithHyperlinks
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a second worksheet
        int detailsIndex = workbook.Worksheets.Add();
        Worksheet summarySheet = workbook.Worksheets[0];
        Worksheet detailsSheet = workbook.Worksheets[detailsIndex];

        // Name the worksheets
        summarySheet.Name = "Summary";
        detailsSheet.Name = "Details";

        // Populate some sample data
        summarySheet.Cells["A1"].PutValue("Summary Data");
        detailsSheet.Cells["A1"].PutValue("Details Data");

        // Define the Excel file name that will be referenced by the hyperlinks
        string excelFileName = "Report.xlsx";

        // Add a hyperlink in each sheet that points to the corresponding sheet in the Excel file
        // Hyperlink format: "Report.xlsx#SheetName!A1"
        summarySheet.Hyperlinks.Add(0, 1, 1, 1, $"{excelFileName}#{summarySheet.Name}!A1");
        detailsSheet.Hyperlinks.Add(0, 1, 1, 1, $"{excelFileName}#{detailsSheet.Name}!A1");

        // Save the workbook as an Excel file (the target of the hyperlinks)
        workbook.Save(excelFileName, SaveFormat.Xlsx);

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true, // retain document structure
            EmbedAttachments = false        // no OLE attachments needed
        };

        // Save the workbook as a PDF; hyperlinks are preserved in the PDF
        workbook.Save("Report.pdf", pdfOptions);
    }
}
