// Title: C# – Export a Worksheet‑Scoped Named Range to PDF with Aspose.Cells
// Description: Shows how to build a workbook, create a range limited to one worksheet (e.g., MyRange covering A1:B3), set it as the print area, and save the result as a PDF containing only that area.
// Keywords: Aspose.Cells | C# | named range PDF | print area | worksheet scoped range | export specific cells | PDF generation | Aspose.Cells .NET | range to PDF | save workbook as PDF
// Common Searches: Aspose.Cells export only named range to PDF | set print area from named range C# | how to save specific cells as PDF using Aspose.Cells | worksheet scoped named range PDF Aspose | C# generate PDF from selected range Aspose.Cells
// Developer Intent: Create a PDF that includes only the cells of a worksheet‑scoped named range.
// Use Cases: Create a catalog page that includes only the product table defined by a named range. | Produce a PDF of an invoice line‑items section without the surrounding worksheet data. | Extract a summary block from a large worksheet into a standalone PDF report.
// AI Prompts: Show C# code to set a worksheet's PrintArea to a named range and save as PDF with Aspose.Cells. | How can I export multiple worksheet‑scoped named ranges to separate PDF files using Aspose.Cells .NET? | Explain the steps to retrieve a named range and use it as the print area for PDF conversion.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPdfNamedRangeDemo
{
    // Shows how to build a workbook, create a range limited to one worksheet (e.g., MyRange covering A1:B3), set it as the print area, and save the result as a PDF containing only that area.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.2);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(0.8);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(2.5);

                // Define a worksheet‑scoped named range that covers A1:B3
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                // RefersTo must include the sheet name and be prefixed with '='
                workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!A1:B3";

                // Retrieve the Range object for the named range
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                AsposeRange range = namedRange.GetRange(); // Resolve ambiguity with alias

                // Set the print area of the worksheet to the address of the named range
                // This ensures that only this area is exported when saving to PDF
                sheet.PageSetup.PrintArea = range.Address; // uses PrintArea property

                // Save the workbook as PDF (lifecycle: save)
                // The PDF will contain only the defined print area
                workbook.Save("NamedRangeOnly.pdf");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
