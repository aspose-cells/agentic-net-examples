// Title: Convert an Aspose.Cells workbook to PDF and set the Subject built‑in document property in C#
// AI Prompts: Generate C# code that uses Aspose.Cells to create a workbook, assign a Subject built‑in document property, and save it as a PDF with standard properties exported. | Show how to configure PdfSaveOptions.CustomPropertiesExport to include built‑in properties when converting an Excel file to PDF with Aspose.Cells.
// Common Searches: Aspose.Cells C# set Subject property before saving workbook as PDF | How to export built‑in document properties to PDF using PdfSaveOptions in Aspose.Cells | C# example converting Excel to PDF with metadata like Subject using Aspose.Cells | PdfSaveOptions CustomPropertiesExport Standard usage in Aspose.Cells conversion
// Tags: Aspose.Cells PDF conversion with built‑in document properties | set workbook Subject property Aspose.Cells | PdfSaveOptions CustomPropertiesExport Standard | export Excel metadata to PDF Aspose.Cells | C# workbook to PDF with subject categorization

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a new workbook, adds sample data, sets the Subject built‑in document property, configures PdfSaveOptions to export standard properties, and saves the workbook as a PDF file.
class ConvertWorkbookToPdfWithSubject
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for PDF conversion");

        // Define the subject property for categorization
        workbook.BuiltInDocumentProperties.Subject = "Financial Report Q1";

        // Configure PDF save options to export built‑in/custom properties
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // Convert the workbook to PDF using the specified options
        workbook.Save("ConvertedWorkbook.pdf", pdfOptions);
    }
}
