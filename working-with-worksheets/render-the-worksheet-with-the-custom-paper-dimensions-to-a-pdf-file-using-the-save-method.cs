// Title: Export a worksheet with Letter-size paper setup to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that configures a worksheet's PageSetup to Letter size and saves the workbook as a PDF with Aspose.Cells. | Show how to set custom paper dimensions in points for a worksheet before calling Workbook.Save to generate a PDF. | Demonstrate adding sample cell data and exporting the worksheet to a PDF file with a specific page size using Aspose.Cells.
// Common Searches: Aspose.Cells C# set worksheet page size to Letter before PDF export | how to define custom paper dimensions in points for PDF output with Aspose.Cells | C# save workbook as PDF with specific page setup using Aspose.Cells | export worksheet to PDF with custom page size using Aspose.Cells .NET
// Tags: worksheet page setup paper size Aspose.Cells | export worksheet to PDF C# Aspose | custom paper dimensions points Aspose.Cells | save workbook as PDF specific page size | Aspose.Cells PDF export page layout

using Aspose.Cells;
using System;

// The example creates a new workbook, sets the first worksheet's page size to Letter (8.5×11 inches), adds sample content, and saves the worksheet as a PDF file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set paper size to Letter (8.5 x 11 inches)
            // 1 point = 1/72 inch, so Letter size matches the desired dimensions.
            worksheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;

            // Optional: add some content to the worksheet
            worksheet.Cells["A1"].PutValue("Sample content on a custom-sized PDF page.");

            // Save the worksheet as a PDF file
            workbook.Save("CustomPaperDimensions.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
