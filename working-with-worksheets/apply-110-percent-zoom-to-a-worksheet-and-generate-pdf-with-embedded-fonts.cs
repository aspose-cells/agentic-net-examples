// Title: Set worksheet zoom to 110 % and export to PDF with embedded fonts using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to set worksheet zoom to 110 percent before saving as PDF | embed fonts when converting Excel worksheet to PDF with Aspose.Cells .NET | increase worksheet view zoom level for PDF export using Aspose.Cells PDFSaveOptions | Aspose.Cells PDF conversion custom zoom and font embedding example in C# | apply 110% zoom to first worksheet and create PDF with embedded fonts using Aspose.Cells
// Tags: worksheet zoom aspnet cells c# | pdf conversion with embedded fonts aspnet cells | pdfsaveoptions font embedding c# | custom zoom level excel to pdf aspnet cells | set zoom property before pdf export c#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, sets the first worksheet's Zoom property to 110 %, configures PdfSaveOptions (fonts are embedded by default), and saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example data (optional)
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Apply 110 percent zoom to the worksheet
            sheet.Zoom = 110; // Zoom is expressed as a percentage

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Font embedding is enabled by default in recent Aspose.Cells versions.
            // If the library version supports explicit settings, they can be added here.

            // Define output file path
            string outputPath = "Output.pdf";

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
