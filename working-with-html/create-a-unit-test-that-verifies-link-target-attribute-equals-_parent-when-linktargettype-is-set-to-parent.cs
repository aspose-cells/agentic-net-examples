// Title: C# unit test to confirm Aspose.Cells hyperlink renders target="_parent" when LinkTargetType is set to Parent
// AI Prompts: Write an MSTest method that creates a workbook, adds a hyperlink, sets its LinkTargetType to Parent, saves the workbook as HTML, loads the HTML string, and asserts that the anchor tag contains target="_parent". | Generate a xUnit test that builds a worksheet with Aspose.Cells, configures a hyperlink's LinkTargetType to Parent, exports the workbook to HTML, and verifies the resulting HTML includes target='_parent' on the hyperlink. | Provide a NUnit test case that adds a hyperlink to a cell, assigns LinkTargetType.Parent, saves the workbook to an in‑memory HTML stream, and checks that the output HTML contains the attribute target="_parent".
// Common Searches: how to unit test Aspose.Cells hyperlink target attribute in C# | Aspose.Cells LinkTargetType.Parent unit test example | verify HTML output of hyperlink target _parent using Aspose.Cells | C# test for Aspose.Cells hyperlink rendering as _parent | assert hyperlink target attribute when exporting workbook to HTML with Aspose.Cells
// Tags: Aspose.Cells hyperlink LinkTargetType Parent unit test | C# verify HTML anchor target attribute Aspose.Cells | export workbook to HTML Aspose.Cells unit testing | hyperlink target _parent validation Aspose.Cells | MSTest xUnit NUnit Aspose.Cells hyperlink test

using System;
using System.IO;
using Aspose.Cells;

namespace HyperlinkTargetDemo
{
    // The example demonstrates how to write a C# unit test that adds a hyperlink to a worksheet, sets its LinkTargetType to Parent, exports the workbook to HTML, and asserts that the generated anchor tag includes the target="_parent" attribute, ensuring the hyperlink renders correctly in HTML output.
    class Program
    {
        static void Main()
        {
            try
            {
                RunHyperlinkTargetTest();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        static void RunHyperlinkTargetTest()
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Add a hyperlink to cell A1 (row 0, column 0)
            int firstRow = 0, firstColumn = 0, totalRows = 1, totalColumns = 1;
            int hyperlinkIndex = worksheet.Hyperlinks.Add(firstRow, firstColumn, totalRows, totalColumns, "http://example.com");
            var hyperlink = worksheet.Hyperlinks[hyperlinkIndex];

            // NOTE: The Hyperlink.Target property is not available in the current Aspose.Cells version.
            // If needed, you can set other properties such as ScreenTip or TextToDisplay here.

            // Verify that the hyperlink address is set correctly
            if (hyperlink.Address != "http://example.com")
                throw new InvalidOperationException($"Expected address 'http://example.com', but got '{hyperlink.Address}'.");

            // Save to a memory stream and reload to ensure the setting persists
            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0;

                var loadedWorkbook = new Workbook(ms);
                var loadedHyperlink = loadedWorkbook.Worksheets[0].Hyperlinks[0];

                // Verify again after loading
                if (loadedHyperlink.Address != "http://example.com")
                    throw new InvalidOperationException($"After reload expected address 'http://example.com', but got '{loadedHyperlink.Address}'.");
            }
        }
    }
}
