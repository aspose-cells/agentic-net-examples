using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsPdfNamedRangeDemo
{
    // Alias to avoid conflict with System.Range
    using AsposeRange = Aspose.Cells.Range;

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (A1:D6)
                for (int row = 0; row < 6; row++)
                {
                    for (int col = 0; col < 4; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Define a named range that covers B2:D5
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                // RefersTo must include the sheet name and start with '='
                namedRange.RefersTo = $"={sheet.Name}!B2:D5";

                // Retrieve the actual Range object from the name
                AsposeRange range = namedRange.GetRange();

                // Set the worksheet's print area to the same address as the named range
                // namedRange.RefersTo format: "=Sheet1!B2:D5"
                // Extract the address part after '!'
                string printArea = namedRange.RefersTo.Split('!')[1].TrimStart('=');
                sheet.PageSetup.PrintArea = printArea;

                // Prepare PDF save options (no ExportArea property for PdfSaveOptions)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF; only the area defined by the print area will be included
                workbook.Save("NamedRangeOnly.pdf", pdfOptions);

                Console.WriteLine("PDF generated with only the named range area.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}