// Title: Export Only a Slicer Region to PDF by Setting Print Area with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to define a worksheet's print area to the slicer range (e.g., C5:F10) using Aspose.Cells, then save the workbook as a PDF so that only the slicer region appears in the output.
// Keywords: Aspose.Cells print area PDF | C# export slicer to PDF | Aspose.Cells set print area | PDFSaveOptions worksheet range | slicer PDF report Aspose.Cells | .NET PDF export specific range | Aspose.Cells worksheet print area
// Common Searches: Aspose.Cells export slicer only to PDF | Set print area for PDF export in C# Aspose.Cells | How to limit PDF output to a worksheet range Aspose.Cells | C# Aspose.Cells PDFSaveOptions print area example | Export dashboard slicer as PDF using Aspose.Cells
// Developer Intent: Create a PDF that contains only the slicer area by defining the worksheet's print area.
// Use Cases: Generate a clean PDF snapshot of a dashboard slicer for client delivery. | Automate reporting where only slicer controls need to be shared, omitting other worksheet data. | Integrate into a CI pipeline to produce printable slicer sections for documentation.
// AI Prompts: Write C# code with Aspose.Cells that sets PageSetup.PrintArea to a slicer range and saves the workbook as a PDF. | Show how to export only the slicer region of a worksheet to PDF using Aspose.Cells .NET. | Explain the steps to configure PdfSaveOptions and PrintArea for limiting PDF output to a specific range.

using System;
using Aspose.Cells;

// Demonstrates how to define a worksheet's print area to the slicer range (e.g., C5:F10) using Aspose.Cells, then save the workbook as a PDF so that only the slicer region appears in the output.
class SlicerPdfReport
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // (Optional) Add some sample data to the sheet
            worksheet.Cells["A1"].PutValue("Sample Data");
            worksheet.Cells["B1"].PutValue(123);
            worksheet.Cells["A2"].PutValue("More Data");
            worksheet.Cells["B2"].PutValue(456);

            // Define the range that contains the slicer (example: C5:F10)
            // This sets the print area so that only this region will be exported to PDF
            worksheet.PageSetup.PrintArea = "C5:F10";

            // Create PDF save options (default options are sufficient for printing the defined area)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file; only the defined print area will appear in the output
            workbook.Save("SlicerReport.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
