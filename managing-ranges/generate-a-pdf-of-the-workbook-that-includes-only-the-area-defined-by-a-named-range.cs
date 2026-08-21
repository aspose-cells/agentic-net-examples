// Title: Export a Named Range to PDF with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, defines a named range (e.g., A1:B3), sets the worksheet's PrintArea to that range, and saves the file as a PDF using PdfSaveOptions. The resulting PDF contains only the cells covered by the named range.
// Keywords: Aspose.Cells C# PDF export | named range to PDF | set print area Aspose.Cells | PdfSaveOptions example | export selected cells as PDF | Aspose.Cells .NET tutorial | C# generate PDF from range | Aspose.Cells Europe | Aspose.Cells United States
// Common Searches: how to export a named range to PDF using Aspose.Cells | C# set worksheet print area for PDF conversion | Aspose.Cells save only selected cells as PDF | PdfSaveOptions named range example | Aspose.Cells generate PDF from specific range
// Developer Intent: Create a PDF that includes only the cells defined by a workbook's named range.
// Use Cases: Produce a compact report PDF that shows just a table header and data rows defined by a named range. | Share a specific section of a large spreadsheet with clients without exposing the entire worksheet. | Automate per‑invoice PDF generation where each invoice area is a named range in the source workbook.
// AI Prompts: Show C# code that assigns a named range to the worksheet PrintArea and saves the workbook as a PDF with Aspose.Cells. | Demonstrate how to customize PdfSaveOptions (page size, orientation, compression) while exporting only a named range. | Explain the steps to retrieve a named range address and use it as the print area before PDF conversion in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsPdfNamedRangeDemo
{
    // Creates a workbook, defines a named range (e.g., A1:B3), sets the worksheet's PrintArea to that range, and saves the file as a PDF using PdfSaveOptions. The resulting PDF contains only the cells covered by the named range.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (A1:B3)
                worksheet.Cells["A1"].PutValue("Header1");
                worksheet.Cells["B1"].PutValue("Header2");
                worksheet.Cells["A2"].PutValue("Item1");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("Item2");
                worksheet.Cells["B3"].PutValue(20);

                // Define a named range that covers the area we want to export
                int nameIndex = workbook.Worksheets.Names.Add("ExportRange");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                // RefersTo must include the sheet name and be prefixed with '='
                namedRange.RefersTo = $"={worksheet.Name}!A1:B3";

                // Retrieve the actual Range object from the named range
                Aspose.Cells.Range range = namedRange.GetRange();

                // Set the worksheet's print area to the address of the named range
                // This ensures that only this area is considered when saving to PDF
                worksheet.PageSetup.PrintArea = range.Address;

                // Create PDF save options (default options are sufficient)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF; only the defined print area will be included
                workbook.Save("ExportedNamedRange.pdf", pdfOptions);

                Console.WriteLine("PDF generated with only the named range area.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
