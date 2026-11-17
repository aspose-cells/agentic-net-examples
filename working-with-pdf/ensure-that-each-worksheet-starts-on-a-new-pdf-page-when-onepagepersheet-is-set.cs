// Title: Save an Aspose.Cells workbook as PDF with each worksheet on a separate page using OnePagePerSheet in C#
// Common Searches: C# Aspose.Cells export each worksheet to its own PDF page | How to configure PDFSaveOptions for separate pages per sheet in Aspose.Cells | Aspose.Cells .NET generate PDF with one page per worksheet | Saving workbook as PDF with per‑sheet pagination using Aspose.Cells | Why PDF output merges worksheets despite PDFSaveOptions settings
// Tags: Aspose.Cells PDFSaveOptions OnePagePerSheet | C# export workbook to PDF per worksheet | Aspose.Cells separate PDF pages for each sheet | OnePagePerSheet PDF conversion .NET | render worksheets on individual PDF pages Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook with three worksheets, sets PdfSaveOptions.OnePagePerSheet = true, and saves the file as a PDF where each worksheet starts on a new page.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "FirstSheet";
        sheet1.Cells["A1"].PutValue("Data in first sheet");
        sheet1.Cells["A2"].PutValue("More data...");

        // Add a second worksheet with its own data
        Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
        sheet2.Cells["A1"].PutValue("Data in second sheet");
        sheet2.Cells["A2"].PutValue("Additional data...");

        // Add a third worksheet
        Worksheet sheet3 = workbook.Worksheets.Add("ThirdSheet");
        sheet3.Cells["A1"].PutValue("Data in third sheet");
        sheet3.Cells["A2"].PutValue("Further data...");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure each worksheet is rendered on a separate PDF page
            OnePagePerSheet = true
        };

        // Save the workbook as a PDF file
        workbook.Save("WorksheetsPerPage.pdf", pdfOptions);
    }
}
