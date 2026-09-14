// Title: How to export an Excel workbook to PDF with all columns forced onto a single page per sheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and saves it as a PDF where each worksheet's columns are compressed onto one page using PdfSaveOptions.AllColumnsInOnePagePerSheet. | Describe the steps to configure Aspose.Cells PdfSaveOptions so that every sheet fits all its columns on a single PDF page, then perform the workbook.Save call in a .NET project. | Provide a C# snippet that exports a workbook to PDF with landscape orientation while keeping the AllColumnsInOnePagePerSheet flag enabled.
// Common Searches: Aspose.Cells .NET export Excel to PDF with all columns on one page per sheet | PdfSaveOptions.AllColumnsInOnePagePerSheet example C# | Fit entire worksheet width onto a single PDF page using Aspose.Cells | How to force single-page PDF output for each sheet in Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions AllColumnsInOnePagePerSheet | C# single-page PDF export per worksheet | fit all worksheet columns onto one PDF page | landscape PDF layout Aspose.Cells | configure PDF save options Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Loads an Excel workbook, sets PdfSaveOptions.AllColumnsInOnePagePerSheet to true so every worksheet's columns are placed on one PDF page, and saves the result as a PDF file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options to fit all columns on one page per sheet
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // This option forces all columns of each worksheet to be placed on a single PDF page
            AllColumnsInOnePagePerSheet = true
        };

        // Export the workbook to PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
