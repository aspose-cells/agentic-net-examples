// Title: Set A4 landscape page size for a worksheet and export to PDF using Aspose.Cells in C#
// AI Prompts: Configure the first worksheet's PageSetup to PaperSizeType.PaperA4 and PageOrientationType.Landscape, then save the workbook as a PDF with Aspose.Cells. | Generate a PDF from an Excel workbook in C# with A4 paper size and landscape orientation using Aspose.Cells.
// Common Searches: how to export Excel to PDF with A4 landscape using Aspose.Cells C# | Aspose.Cells set worksheet page orientation to landscape before PDF conversion | C# Aspose.Cells change PDF paper size to A4 | save workbook as PDF with custom page setup Aspose.Cells .NET | Aspose.Cells PDF export page setup orientation paper size
// Tags: Aspose.Cells set worksheet PDF page size | Aspose.Cells PDF landscape orientation | C# Aspose.Cells export workbook to PDF | Aspose.Cells configure page setup for PDF | PaperSizeType PaperA4 Aspose.Cells | PageOrientationType Landscape PDF export

using Aspose.Cells;
using System.Drawing.Printing;

// Load an existing workbook (replace with your file path)
Workbook workbook = new Workbook("input.xlsx");

// Access the first worksheet (or any specific worksheet)
Worksheet sheet = workbook.Worksheets[0];

// Configure PDF page settings
sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;          // Set page size to A4
sheet.PageSetup.Orientation = PageOrientationType.Landscape; // Set orientation to landscape

// Save the workbook as a PDF with the configured page settings
workbook.Save("output.pdf", SaveFormat.Pdf);
