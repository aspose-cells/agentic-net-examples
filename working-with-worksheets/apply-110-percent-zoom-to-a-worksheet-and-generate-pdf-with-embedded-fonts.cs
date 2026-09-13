// Title: Apply 110% zoom to an Excel worksheet and export to PDF with embedded fonts using Aspose.Cells for .NET
// AI Prompts: Create a new Workbook, set the first Worksheet's Zoom property to 110, and save it as a PDF while ensuring fonts are embedded with Aspose.Cells in C#. | Using Aspise.Cells, adjust the worksheet zoom level to 110% and generate a PDF output that includes embedded fonts, handling any exceptions that may occur.
// Common Searches: Aspose.Cells set worksheet zoom to 110 percent before PDF conversion C# | how to export Excel to PDF with custom zoom using Aspose.Cells .NET | C# Aspose.Cells PDF save options embed fonts automatically | increase worksheet view scale and generate PDF with Aspose.Cells | apply zoom level to worksheet and save as PDF in Aspose.Cells example
// Tags: worksheet zoom property Aspose.Cells | pdf export with font embedding Aspose.Cells | set zoom before pdf conversion C# | Aspose.Cells PdfSaveOptions usage | excel to pdf with custom zoom .NET

using Aspose.Cells;
using System;

// // This program creates a new workbook, sets the first worksheet's Zoom to 110%, and saves the workbook as a PDF using Aspose.Cells with default PDF save options that embed fonts.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Apply 110 percent zoom to the worksheet
            sheet.Zoom = 110;

            // Configure PDF save options (no need to embed standard fonts for Aspose.Cells)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF with the specified options
            workbook.Save("output.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
