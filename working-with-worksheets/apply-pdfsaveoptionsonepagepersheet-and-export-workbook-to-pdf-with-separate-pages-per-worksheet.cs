// Title: Generate a PDF from an Aspose.Cells workbook with one PDF page per worksheet using C#
// AI Prompts: Provide C# code that loads an existing Excel file and saves it as a PDF where each worksheet appears on its own page using Aspose.Cells PdfSaveOptions. | Demonstrate how to configure Aspose.Cells PdfSaveOptions.OnePagePerSheet in C# before calling Workbook.Save. | Show how to export a newly created workbook to PDF with separate pages per sheet using Aspose.Cells in .NET.
// Common Searches: Aspose.Cells C# how to export each worksheet to a separate PDF page | PdfSaveOptions OnePagePerSheet property usage example .NET | Save Excel workbook as PDF with one page per sheet using Aspose.Cells library | C# code sample for PDF export with one page per worksheet in Aspose.Cells | Aspose.Cells PDF save options to generate individual pages per sheet
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | C# export workbook to PDF per worksheet | PDF save options separate sheet pages Aspose | .NET generate PDF with one page per Excel sheet | Aspose.Cells PDF export per sheet

// Create a new workbook (or load an existing one)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(); // create rule

// If you need to load an existing workbook, uncomment the line below and provide the path
// workbook = new Aspose.Cells.Workbook("input.xlsx"); // load rule

// Configure PDF save options to generate one PDF page per worksheet
Aspose.Cells.PdfSaveOptions pdfOptions = new Aspose.Cells.PdfSaveOptions();
pdfOptions.OnePagePerSheet = true; // set the required option

// Save the workbook to PDF using the configured options
workbook.Save("output.pdf", pdfOptions); // save rule
