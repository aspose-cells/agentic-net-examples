// Title: Set a Print Area in Aspose.Cells .NET (C#) and Export Only Selected Cells to PDF
// Description: C# example that creates a workbook, fills cells A1:B4, defines the print area as $A$1:$B$3, optionally fits the area to one page, and saves the file as PDF so that only the specified range appears in the output.
// Keywords: Aspose.Cells | C# | PrintArea | PDF export | named range | PageSetup.PrintArea | FitToPagesWide | FitToPagesTall | GitHub example | Excel to PDF
// Common Searches: Aspose.Cells set print area before PDF conversion | limit PDF output to specific cells C# | PageSetup.PrintArea Aspose.Cells example | export selected range to PDF Aspose.Cells .NET | fit print area to one PDF page C#
// Developer Intent: Define a worksheet print area so that only the chosen cells are included when the workbook is saved as a PDF.
// Use Cases: Create a compact PDF containing only a header and a few data rows. | Generate PDFs where the printable region is controlled programmatically without opening Excel. | Produce one‑page reports by fitting a custom print area to a single PDF page.
// AI Prompts: Show C# code that assigns a named range as the print area in Aspose.Cells before saving to PDF. | How can I use PageSetup.FitToPagesWide and FitToPagesTall to fit a print area onto one PDF page in Aspose.Cells? | Explain how to verify that the defined print area is applied during PDF conversion with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaPdfDemo
{
    // C# example that creates a workbook, fills cells A1:B4, defines the print area as $A$1:$B$3, optionally fits the area to one page, and saves the file as PDF so that only the specified range appears in the output.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Laptop");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["A3"].PutValue("Phone");
                sheet.Cells["B3"].PutValue(800);
                sheet.Cells["A4"].PutValue("Tablet");
                sheet.Cells["B4"].PutValue(500);

                // Define the print area directly (covers A1:B3)
                sheet.PageSetup.PrintArea = "$A$1:$B$3";

                // Optional: adjust page setup to fit the print area on a single page
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = 1;

                // Save the workbook as PDF. The defined print area limits the PDF content.
                string outputPath = "PrintAreaLimited.pdf";
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"PDF generated with print area limited to the range A1:B3: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
