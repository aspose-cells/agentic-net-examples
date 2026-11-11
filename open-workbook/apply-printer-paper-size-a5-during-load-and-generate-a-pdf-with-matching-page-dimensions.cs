// Title: Apply A5 paper size to every worksheet while loading an Excel file and export it as a PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set A5 page size for all worksheets before saving to PDF | How to change paper size to A5 in each worksheet using Aspose.Cells | Export Excel to PDF with A5 dimensions using Aspose.Cells .NET | Set worksheet PageSetup PaperSize to A5 programmatically in C# | Load workbook and apply A5 printer paper size with Aspose.Cells
// Tags: worksheet pagesetup a5 format | aspose.cells pdf export a5 page dimensions | c# configure workbook paper size a5 | load excel set a5 pagesetup aspnet | excel to pdf a5 format conversion

using Aspose.Cells;
using System;

// The example loads an existing Excel workbook, sets each worksheet's PageSetup.PaperSize to PaperA5, and saves the workbook as a PDF, producing PDF pages that match A5 dimensions.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Apply A5 paper size to each worksheet's page setup
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA5;
        }

        // Save the workbook as PDF; the PDF pages will match the A5 dimensions
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
