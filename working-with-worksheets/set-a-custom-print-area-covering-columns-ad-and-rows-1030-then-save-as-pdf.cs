// Title: Set a custom print area (A10:D30) on a worksheet and export it to PDF with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set print area for specific rows and columns before PDF export | How to define a print range A10 to D30 and save as PDF using Aspose.Cells .NET | Export only selected cells to PDF with Aspose.Cells workbook in C# | C# code to configure page setup print area and generate PDF via Aspose.Cells | Saving a worksheet with custom print area to PDF in Aspose.Cells for .NET
// Tags: print area configuration Aspose.Cells | range-based PDF export Aspose.Cells | worksheet page setup to PDF C# | custom print range PDF conversion .NET | Aspose.Cells workbook save as PDF

using System;
using Aspose.Cells;

// The example creates a new workbook, accesses the first worksheet, sets its print area to cells A10 through D30, and then saves the workbook directly as a PDF file named "CustomPrintArea.pdf" using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the custom print area (columns A‑D, rows 10‑30)
        sheet.PageSetup.PrintArea = "A10:D30";

        // Save the workbook as a PDF file
        workbook.Save("CustomPrintArea.pdf", SaveFormat.Pdf);
    }
}
