// Title: How to set a custom print area (A10:D30) and save a worksheet as PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set print area A10:D30 before PDF conversion | C# code to define custom print range and export to PDF using Aspose.Cells | Save specific rows and columns as PDF with Aspose.Cells page setup | Fit printed area to one page width when generating PDF with Aspose.Cells | Create PDF from worksheet with limited print area using Aspose.Cells .NET
// Tags: custom print area Aspose.Cells | worksheet to PDF conversion C# | PageSetup.PrintArea range Aspose.Cells | fit to page width Aspose.Cells PDF | export selected cells as PDF Aspose.Cells | Aspose.Cells PDF save format .NET

using Aspose.Cells;
using System;

// The example creates a new Workbook, accesses the first Worksheet, sets the PageSetup.PrintArea to "A10:D30", configures the page to fit one page wide, and saves the workbook as a PDF file named "CustomPrintArea.pdf" using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the custom print area to columns A‑D and rows 10‑30
        sheet.PageSetup.PrintArea = "A10:D30";

        // Optional: fit the print area to one page width
        sheet.PageSetup.FitToPagesWide = 1;
        sheet.PageSetup.FitToPagesTall = 0;

        // Save the workbook as a PDF file
        workbook.Save("CustomPrintArea.pdf", SaveFormat.Pdf);
    }
}
