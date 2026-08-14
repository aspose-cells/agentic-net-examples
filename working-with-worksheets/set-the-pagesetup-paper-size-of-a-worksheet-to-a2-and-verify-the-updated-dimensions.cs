// Title: C# – Set Worksheet PageSetup PaperSize to A2 and Verify Dimensions with Aspose.Cells
// Description: Demonstrates how to create a workbook, set the first worksheet's PageSetup.PaperSize to PaperSizeType.PaperA2, read the resulting PaperWidth and PaperHeight in inches, render the sheet to obtain the actual page size, confirm a page is generated, and save the file as A2PaperSize.xlsx.
// Keywords: Aspose.Cells C# | set worksheet paper size A2 | PageSetup PaperSize | PaperWidth PaperHeight inches | SheetRender GetPageSizeInch | verify page dimensions | export A2 workbook | large‑format printing | Aspose.Cells example
// Common Searches: Aspose.Cells set A2 paper size .NET | how to get worksheet dimensions after setting paper size | SheetRender GetPageSizeInch example | C# verify A2 page width and height in Aspose.Cells | check page count after changing paper size
// Developer Intent: Configure a worksheet for A2 printing and confirm that the physical and rendered dimensions match the A2 standard.
// Use Cases: Prepare reports or posters that require A2 paper size before exporting to PDF or image formats. | Validate paper‑size settings in automated document‑generation pipelines. | Generate a preview of the first printed page to ensure layout fits A2 dimensions.
// AI Prompts: Generate C# code that sets a worksheet's PaperSetup.PaperSize to PaperSizeType.PaperA2 and logs PaperWidth and PaperHeight. | Explain how SheetRender.GetPageSizeInch can be used to retrieve the rendered page size after configuring A2 paper size in Aspose.Cells. | Write a method that compares the rendered page size (in inches) with the standard A2 dimensions (16.54 × 23.39) and throws an error if they differ.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook, set the first worksheet's PageSetup.PaperSize to PaperSizeType.PaperA2, read the resulting PaperWidth and PaperHeight in inches, render the sheet to obtain the actual page size, confirm a page is generated, and save the file as A2PaperSize.xlsx.
class SetPaperSizeA2
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the paper size of the worksheet to A2
            worksheet.PageSetup.PaperSize = PaperSizeType.PaperA2;

            // Verify the paper size enum value
            Console.WriteLine("PaperSize set to: " + worksheet.PageSetup.PaperSize);

            // Verify the physical dimensions (in inches) after setting the size
            Console.WriteLine("Paper Width (inches): " + worksheet.PageSetup.PaperWidth);
            Console.WriteLine("Paper Height (inches): " + worksheet.PageSetup.PaperHeight);

            // Render the worksheet to obtain page size information
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            SheetRender sheetRender = new SheetRender(worksheet, options);

            // Ensure at least one page exists before querying its size
            if (sheetRender.PageCount > 0)
            {
                float[] pageSize = sheetRender.GetPageSizeInch(0);
                Console.WriteLine($"Rendered page size (inches): Width={pageSize[0]}, Height={pageSize[1]}");
            }
            else
            {
                Console.WriteLine("No pages were generated for the worksheet.");
            }

            // Save the workbook
            string outputPath = "A2PaperSize.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
