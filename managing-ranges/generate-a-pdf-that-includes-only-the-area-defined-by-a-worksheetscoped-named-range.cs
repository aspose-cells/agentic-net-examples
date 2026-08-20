// Title: Export a Worksheet‑Scoped Named Range to PDF with Aspose.Cells for .NET (C#)
// Description: This example shows how to create a workbook, define a worksheet‑scoped named range (B2:C4), set the worksheet's print area to that range, and save the result as a PDF so that only the named range appears in the output file.
// Keywords: Aspose.Cells | C# | .NET | PDF export | named range | worksheet scoped range | PrintArea | PdfSaveOptions | export specific range to PDF | Aspose.Cells example
// Common Searches: Aspose.Cells export only a named range to PDF | Set print area to a named range in C# | Worksheet‑scoped named range PDF export .NET | How to limit PDF output to a cell range using Aspose.Cells | C# code for PDFSaveOptions with named range
// Developer Intent: Generate a PDF that contains exclusively the cells defined by a worksheet‑scoped named range.
// Use Cases: Create a PDF report that includes just a data table defined by a named range, hiding other worksheet content. | Produce individual invoice PDFs by assigning each invoice area a named range and exporting it separately. | Offer a preview of a selected chart or table by setting its named range as the print area before PDF conversion.
// AI Prompts: Show C# code to set a worksheet‑scoped named range as the print area and export it to PDF with Aspose.Cells. | How can I export only the cells of a named range to a PDF using Aspose.Cells for .NET? | Explain the steps to create a named range, register it, retrieve the Range object, and limit PDF output to that area.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPdfNamedRangeDemo
{
    // This example shows how to create a workbook, define a worksheet‑scoped named range (B2:C4), set the worksheet's print area to that range, and save the result as a PDF so that only the named range appears in the output file.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["C1"].PutValue("Header3");
                sheet.Cells["A2"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);
                sheet.Cells["C2"].PutValue(30);
                sheet.Cells["A3"].PutValue(40);
                sheet.Cells["B3"].PutValue(50);
                sheet.Cells["C3"].PutValue(60);
                sheet.Cells["A4"].PutValue(70);
                sheet.Cells["B4"].PutValue(80);
                sheet.Cells["C4"].PutValue(90);

                // ---------- Define a worksheet‑scoped named range ----------
                // Create a range B2:C4 and assign a name to it
                AsposeRange namedRange = sheet.Cells.CreateRange("B2", "C4");
                namedRange.Name = "MyRange";

                // Register the name in the worksheet's name collection (worksheet‑scoped)
                int nameIdx = sheet.Workbook.Worksheets.Names.Add("MyRange");
                // RefersTo must include the sheet name and start with '='
                sheet.Workbook.Worksheets.Names[nameIdx].RefersTo = $"={sheet.Name}!B2:C4";

                // ---------- Retrieve the range via the Name object ----------
                Name nameObj = sheet.Workbook.Worksheets.Names[nameIdx];
                AsposeRange range = nameObj.GetRange(); // obtains the actual Range object

                // ---------- Set the print area to the named range ----------
                // PrintArea expects an address without the leading '='
                sheet.PageSetup.PrintArea = range.RefersTo?.TrimStart('=');

                // ---------- Save the workbook as PDF ----------
                // The defined print area will be respected during PDF export.
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                workbook.Save("NamedRangeArea.pdf", pdfOptions);

                Console.WriteLine("PDF generated successfully with only the named range area.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
