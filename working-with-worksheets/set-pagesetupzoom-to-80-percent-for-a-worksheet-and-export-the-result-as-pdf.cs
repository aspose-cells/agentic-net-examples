// Title: Set worksheet print zoom to 80 % and export to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set worksheet print zoom to 80 percent before PDF conversion | How to change PageSetup.Zoom to 80% and save workbook as PDF in .NET | C# export Excel worksheet to PDF with custom page scaling using Aspose.Cells | Adjust worksheet print scaling to 80% with Aspose.Cells and generate PDF file | Set page zoom on a worksheet and convert to PDF using Aspose.Cells API
// Tags: PageSetup.Zoom property C# | worksheet print scaling Aspose.Cells | PDF export with custom zoom Aspose.Cells | set worksheet zoom before PDF save | Aspose.Cells page scaling for PDF conversion

using Aspose.Cells;
using System;

// The example creates a new workbook, accesses the first worksheet, sets its PageSetup.Zoom to 80 percent, and then saves the workbook as a PDF file named output.pdf.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the page zoom to 80%
        sheet.PageSetup.Zoom = 80;

        // Export the workbook (first worksheet) to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
