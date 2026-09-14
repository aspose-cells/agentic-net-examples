// Title: How to limit Aspose.Cells PDF conversion to the first 10 pages using PdfSaveOptions in C#
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells and saves it as a PDF containing only the first ten pages by configuring PdfSaveOptions. | Show how to set PdfSaveOptions.PageCount to 10 before calling Workbook.Save to export a PDF with a restricted page count in .NET.
// Common Searches: Aspose.Cells C# export first 10 pages of workbook to PDF | PdfSaveOptions.PageCount property usage example | How to save only a subset of Excel pages as PDF with Aspose.Cells | Limit PDF output pages when converting Excel to PDF in .NET | C# code sample for restricting PDF page count with Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions page count | C# export Excel to PDF limited pages | Aspose.Cells limit PDF pages | PdfSaveOptions PageCount C# | Excel to PDF page range Aspose

using Aspose.Cells;

// Load an existing workbook (replace with your actual file path)
Workbook workbook = new Workbook("input.xlsx");

// Create PDF save options and limit the output to the first ten pages
PdfSaveOptions pdfOptions = new PdfSaveOptions();
pdfOptions.PageCount = 10; // limit to first 10 pages

// Save the workbook as PDF using the configured options
workbook.Save("output.pdf", pdfOptions);
