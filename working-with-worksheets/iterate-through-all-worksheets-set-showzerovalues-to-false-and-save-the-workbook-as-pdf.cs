// Title: Hide Zero Values Across All Worksheets and Save as PDF Using Aspose.Cells for .NET (C#)
// Description: A C# sample that creates or loads an Aspose.Cells workbook, iterates through each worksheet, sets the DisplayZeros property to false to suppress zero cells, and then exports the whole workbook to a PDF file with PdfSaveOptions.
// Keywords: Aspose.Cells C# | DisplayZeros false | hide zero values | iterate worksheets | PDF export | PdfSaveOptions | Workbook.Save PDF | suppress zero cells | Aspose.Cells .NET | Excel to PDF conversion
// Common Searches: Aspose.Cells hide zero values C# | Set DisplayZeros false for all worksheets | Export Aspose.Cells workbook to PDF | Iterate worksheets Aspose.Cells example | C# code to suppress zeros in Excel PDF | PdfSaveOptions hide zeros Aspose
// Developer Intent: Disable zero display on every worksheet and generate a PDF from the workbook.
// Use Cases: Financial statements where zero amounts should not appear in the printed PDF. | Invoice batches that need clean PDFs without placeholder zeros across multiple sheets. | Automated reporting pipelines that hide zero values before archiving workbooks as PDFs.
// AI Prompts: Show a C# snippet that loops through all worksheets in an Aspose.Cells workbook, sets DisplayZeros to false, and saves the file as PDF using PdfSaveOptions. | Explain how the DisplayZeros property affects PDF rendering in Aspose.Cells and whether additional PDF options are required. | Modify the example to load an existing workbook from a given path, hide zeros on every sheet, and export it to a PDF file.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsShowZeroExample
{
    // A C# sample that creates or loads an Aspose.Cells workbook, iterates through each worksheet, sets the DisplayZeros property to false to suppress zero cells, and then exports the whole workbook to a PDF file with PdfSaveOptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add sample data with zero values to demonstrate the effect
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            sheet1.Cells["A1"].PutValue(0);
            sheet1.Cells["A2"].PutValue(123);

            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["B1"].PutValue(0);
            sheet2.Cells["B2"].PutValue(456);

            // Iterate through all worksheets and hide zero values
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.DisplayZeros = false; // Do not display zero values
            }

            // Save the workbook as PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}
