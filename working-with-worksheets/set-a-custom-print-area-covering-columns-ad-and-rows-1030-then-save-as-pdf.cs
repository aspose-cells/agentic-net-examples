// Title: Define a custom print area (A10:D30) in Aspose.Cells for .NET and export the worksheet to PDF
// AI Prompts: Set the worksheet's PageSetup.PrintArea to "A10:D30" and save the workbook as a PDF using Aspose.Cells in C#. | Create a new workbook, configure a print range covering columns A through D and rows 10 to 30, then generate a PDF file.
// Common Searches: Aspose.Cells C# set print area A10:D30 before PDF conversion | How to export only a specific range to PDF with Aspose.Cells .NET | C# Aspose.Cells print area for PDF output limited to rows 10-30 | Saving a workbook as PDF with custom print range using Aspose.Cells
// Tags: Aspose.Cells set print area | PageSetup print area PDF conversion | custom range PDF export Aspose.Cells | C# export worksheet to PDF | print area rows 10-30 Aspose.Cells

using System;
using Aspose.Cells;

// // Creates a workbook, sets the print area to cells A10 through D30 on the first worksheet, and saves the file as CustomPrintArea.pdf.
class Program
{
    static void Main()
    {
        // Create a new workbook (you can also load an existing file with Workbook workbook = new Workbook("input.xlsx");)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the custom print area to columns A‑D and rows 10‑30
        sheet.PageSetup.PrintArea = "A10:D30";

        // Save the workbook as a PDF file
        workbook.Save("CustomPrintArea.pdf", SaveFormat.Pdf);
    }
}
