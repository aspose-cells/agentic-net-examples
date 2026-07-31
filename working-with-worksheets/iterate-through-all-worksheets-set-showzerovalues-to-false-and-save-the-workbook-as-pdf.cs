// Title: Hide Zero Values Across All Worksheets and Export Workbook to PDF with Aspose.Cells for .NET (C#)
// Description: Demonstrates creating or loading an Aspose.Cells workbook, populating cells, looping through every worksheet to set DisplayZeros = false (disabling zero display), and saving the entire workbook as a PDF using PdfSaveOptions.
// Keywords: Aspose.Cells | C# | .NET | DisplayZeros | hide zero values | PDF export | iterate worksheets | Workbook to PDF | zero suppression | ShowZeroValues false
// Common Searches: Aspose.Cells hide zeros in all sheets | Set DisplayZeros false for every worksheet C# | Export workbook to PDF after disabling zero display | How to suppress zero values in PDF output Aspose.Cells | Iterate worksheets Aspose.Cells .NET PDF conversion
// Developer Intent: Turn off zero‑value display for every worksheet and generate a PDF file.
// Use Cases: Financial statements where zero amounts should be omitted before printing. | Invoice batches that must hide empty price cells in the final PDF. | Automated report pipelines that apply a global zero‑suppression rule across multiple sheets.
// AI Prompts: Generate C# code with Aspose.Cells that loops through all worksheets, sets DisplayZeros to false, and saves the workbook as a PDF. | Show an example of disabling zero values on newly added worksheets before exporting to PDF using Aspose.Cells. | Explain the impact of the DisplayZeros property on PDF rendering and how to apply it workbook‑wide in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates creating or loading an Aspose.Cells workbook, populating cells, looping through every worksheet to set DisplayZeros = false (disabling zero display), and saving the entire workbook as a PDF using PdfSaveOptions.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Sample data with zero values
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Cells["A1"].PutValue(0);
        sheet1.Cells["A2"].PutValue(123);

        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["B1"].PutValue(0);

        // Iterate through all worksheets and hide zero values
        foreach (Worksheet ws in workbook.Worksheets)
        {
            ws.DisplayZeros = false; // ShowZeroValues = false
        }

        // Save the workbook as PDF
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("output.pdf", pdfOptions);
    }
}
