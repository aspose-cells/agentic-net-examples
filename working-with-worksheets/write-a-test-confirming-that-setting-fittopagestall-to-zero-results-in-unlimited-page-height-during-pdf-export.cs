// Title: C# test: FitToPagesTall = 0 produces unlimited page height in Aspose.Cells PDF export
// Description: Creates a workbook with 200 rows, sets PageSetup.FitToPagesWide = 1 and FitToPagesTall = 0, renders the sheet to count pages, then changes FitToPagesTall to 1 and verifies that the unlimited setting generates more pages before saving as PDF.
// Keywords: Aspose.Cells | FitToPagesTall | C# PDF export | unlimited page height | SheetRender page count | PdfSaveOptions | unit test | page setup pagination
// Common Searches: Aspose.Cells FitToPagesTall zero test | verify unlimited page height PDF Aspose.Cells | C# unit test FitToPagesTall 0 vs 1 | how to count pages with SheetRender Aspose.Cells | PDF pagination when FitToPagesTall is zero
// Developer Intent: Confirm that setting PageSetup.FitToPagesTall to 0 allows a worksheet to span an unlimited number of pages during PDF export.
// Use Cases: Automated regression test to protect the unlimited‑height behavior in future Aspose.Cells releases. | Generating large, vertically unbounded PDF reports from Excel data. | Comparing pagination results between limited (FitToPagesTall = 1) and unlimited (FitToPagesTall = 0) configurations.
// AI Prompts: Write an xUnit test that asserts page count increases when FitToPagesTall is set to 0 versus 1 using Aspose.Cells for .NET. | Explain how the FitToPagesTall property affects PDF pagination in Aspose.Cells C#. | Provide a concise code example that saves a worksheet as a PDF with unlimited height and validates the page count programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a workbook with 200 rows, sets PageSetup.FitToPagesWide = 1 and FitToPagesTall = 0, renders the sheet to count pages, then changes FitToPagesTall to 1 and verifies that the unlimited setting generates more pages before saving as PDF.
class FitToPagesTallTest
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate enough rows to require more than one printed page
        for (int i = 0; i < 200; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Configure page setup: fit all columns to one page wide,
        // and set FitToPagesTall to 0 for unlimited page height
        PageSetup pageSetup = worksheet.PageSetup;
        pageSetup.FitToPagesWide = 1;
        pageSetup.FitToPagesTall = 0; // unlimited height

        // Render the sheet to determine how many pages are generated
        ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();
        SheetRender renderUnlimited = new SheetRender(worksheet, renderOptions);
        int pageCountUnlimited = renderUnlimited.PageCount;

        // Change FitToPagesTall to 1 (limit to a single page tall) and re‑render
        pageSetup.FitToPagesTall = 1;
        SheetRender renderLimited = new SheetRender(worksheet, renderOptions);
        int pageCountLimited = renderLimited.PageCount;

        // Verify that unlimited height produces more pages than the limited case
        if (pageCountUnlimited <= pageCountLimited)
        {
            throw new Exception($"Test failed: unlimited page count ({pageCountUnlimited}) is not greater than limited page count ({pageCountLimited}).");
        }

        // Save the workbook as PDF (the actual PDF content is not examined in this test)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("FitToPagesTall_Unlimited.pdf", pdfOptions);

        Console.WriteLine($"Test passed. Unlimited page count: {pageCountUnlimited}, Limited page count: {pageCountLimited}");
    }
}
