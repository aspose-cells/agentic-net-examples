// Title: Set A4 paper size and landscape orientation for PDF export using Aspose.Cells in C#
// Common Searches: how to export a worksheet to PDF with A4 landscape using Aspose.Cells C# | Aspose.Cells set page setup to A4 and landscape before saving as PDF | C# Aspose.Cells PDF export custom paper size and orientation | change PDF page orientation to landscape in Aspose.Cells workbook | set PaperSizeType.PaperA4 and PageOrientationType.Landscape for PDF output Aspose.Cells
// Tags: Aspose.Cells PDF page size A4 | Aspose.Cells PDF landscape orientation | C# PageSetup PaperSizeType A4 usage | C# PageSetup orientation landscape usage | Aspose.Cells workbook export to PDF with custom layout | Aspose.Cells configure PDF page settings programmatically

using System;
using Aspose.Cells;

// The example creates a Workbook, iterates through each Worksheet, sets PageSetup.PaperSize to PaperA4 and PageSetup.Orientation to Landscape, then saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Iterate through all worksheets to set PDF page settings
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Set the paper size to A4
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Set the orientation to landscape
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;
        }

        // Save the workbook as a PDF with the configured page settings
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
