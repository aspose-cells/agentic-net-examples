// Title: Export a Named Range to PDF with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, defines a named range (DataRange) covering cells A1:B4, sets the worksheet's PrintArea to that range, and saves the workbook as a PDF so that only the named‑range area appears in the output file.
// Keywords: Aspose.Cells | C# | named range PDF | export named range to PDF | print area | Workbook.Save PDF | Aspose.Cells .NET | Excel to PDF | range export | Aspose.Cells API
// Common Searches: Aspose.Cells export named range to PDF C# | set print area from named range Aspose.Cells | save only selected cells as PDF using Aspose.Cells | C# code to generate PDF from a named range | Aspose.Cells limit PDF output to a range
// Developer Intent: Generate a PDF that contains only the cells defined by a named range.
// Use Cases: Create a PDF of a sales table without extra worksheet data by assigning the table’s named range as the print area. | Export just the invoice section of a workbook to PDF, ensuring confidential worksheet content is omitted. | Produce a PDF snippet for a report by setting a predefined named range as the printable area before saving.
// AI Prompts: Show C# code that sets a named range as the PrintArea and exports the workbook to PDF with Aspose.Cells. | How can I export multiple named ranges to separate PDF files using Aspose.Cells for .NET? | Explain how to retrieve a named range address and use it to limit PDF output in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfNamedRangeDemo
{
    // This example creates a workbook, defines a named range (DataRange) covering cells A1:B4, sets the worksheet's PrintArea to that range, and saves the workbook as a PDF so that only the named‑range area appears in the output file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Qty");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(30);

                // Define a named range that covers the data we want in the PDF (A1:B4)
                int nameIndex = workbook.Worksheets.Names.Add("DataRange");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                // RefersTo must start with '=' and include the sheet name
                namedRange.RefersTo = $"={sheet.Name}!A1:B4";

                // Retrieve the actual Range object from the named range (use fully qualified type to avoid ambiguity)
                Aspose.Cells.Range range = namedRange.GetRange();

                // Set the worksheet's print area to the address of the named range
                // This ensures that only this area is considered when printing/saving to PDF
                sheet.PageSetup.PrintArea = range.Address;

                // Save the workbook as PDF. The print area set above limits the output to the named range.
                workbook.Save("NamedRangeOutput.pdf");

                Console.WriteLine("PDF generated with only the named range area.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
