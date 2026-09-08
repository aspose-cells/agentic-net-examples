// Title: Save an Excel workbook to PDF with each worksheet on a separate page using Aspose.Cells PdfSaveOptions.OnePagePerSheet (C#)
// AI Prompts: Write C# code that loads an .xlsx file, sets PdfSaveOptions.OnePagePerSheet to true, and saves it as a PDF where each worksheet starts on its own page. | Show how to configure Aspose.Cells PDF save options to force a new PDF page for every worksheet during conversion in .NET. | Provide a minimal example that creates a workbook, applies OnePagePerSheet, and outputs a multi‑page PDF using Aspose.Cells.
// Common Searches: Aspose.Cells how to export each Excel sheet to a separate PDF page in C# | C# PdfSaveOptions OnePagePerSheet property usage example | Convert workbook to PDF with one page per worksheet using Aspose.Cells .NET | Save Excel file as PDF with each worksheet on its own page Aspose.Cells | Aspose.Cells PDF conversion separate pages per sheet tutorial
// Tags: Aspose.Cells per-sheet PDF export | C# Excel to PDF separate pages | PDF save options for separate worksheets | export workbook as multi‑page PDF | Aspose.Cells PDF conversion per worksheet

// Create a new workbook (or load an existing one)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// If you need to load an existing workbook, uncomment the line below and provide the file path
// Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Configure PDF save options to start each worksheet on a new page
Aspose.Cells.PdfSaveOptions pdfOptions = new Aspose.Cells.PdfSaveOptions
{
    // Ensures that each worksheet is rendered on a separate PDF page
    OnePagePerSheet = true
};

// Save the workbook as a PDF with the specified options
workbook.Save("output.pdf", pdfOptions);
