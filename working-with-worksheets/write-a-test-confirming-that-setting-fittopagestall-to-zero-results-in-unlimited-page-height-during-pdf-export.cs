// Title: C# test to confirm that setting Worksheet.PageSetup.FitToPagesTall to 0 creates unlimited page height in PDF export using Aspose.Cells
// Common Searches: c# Aspose.Cells verify unlimited page height by setting FitToPagesTall to 0 during PDF export | how to check that FitToPagesTall = 0 generates multiple PDF pages with Aspose.Cells | Aspose.Cells PDF export page count test for FitToPagesTall zero in C# | count pages in PDF stream after setting FitToPagesWide and FitToPagesTall with Aspose.Cells
// Tags: worksheet pagesetup fit-to-pages-tall zero unlimited height | aspose.cells pdf export page count verification | c# memorystream pdf page detection aspose.cells | fit-to-pages-wide one page width aspose.cells | automated test for pdf pagination aspose.cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example creates a workbook, fills 200 rows, sets PageSetup.FitToPagesTall to 0 (unlimited height) and FitToPagesWide to 1 (single‑page width), saves the workbook as PDF to a MemoryStream, and counts the "/Type /Page" markers in the PDF content. A page count greater than one confirms that the unlimited page height setting works.
class FitToPagesTallTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with enough rows to require multiple pages when not limited
            for (int row = 0; row < 200; row++)
            {
                sheet.Cells[row, 0].PutValue($"Row {row + 1}");
            }

            // Set FitToPagesTall to 0 (unlimited height) and FitToPagesWide to 1 (fit width to one page)
            sheet.PageSetup.FitToPagesTall = 0;
            sheet.PageSetup.FitToPagesWide = 1;

            // Export the workbook to PDF using a memory stream
            using (MemoryStream pdfStream = new MemoryStream())
            {
                workbook.Save(pdfStream, SaveFormat.Pdf);
                pdfStream.Position = 0; // Reset stream position for reading

                // Simple page count by counting "/Type /Page" occurrences in the PDF content
                string pdfContent = Encoding.ASCII.GetString(pdfStream.ToArray());
                int pageCount = CountOccurrences(pdfContent, "/Type /Page");

                // Verify that more than one page was created, indicating unlimited page height
                if (pageCount > 1)
                {
                    Console.WriteLine("Test passed: FitToPagesTall = 0 results in multiple PDF pages (unlimited height).");
                }
                else
                {
                    Console.WriteLine($"Test failed: Expected multiple pages, but PDF contains {pageCount} page(s).");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to count occurrences of a substring within a string
    private static int CountOccurrences(string source, string substring)
    {
        if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(substring))
            return 0;

        int count = 0;
        int index = 0;
        while ((index = source.IndexOf(substring, index, StringComparison.Ordinal)) != -1)
        {
            count++;
            index += substring.Length;
        }
        return count;
    }
}
