// Title: C# unit test to verify Aspose.Cells HTML export includes gridline borders only when ExportGridLines is enabled
// AI Prompts: Generate a C# NUnit test that creates a workbook, saves it to HTML with HtmlSaveOptions.ExportGridLines set to true, and asserts that the resulting HTML contains a border style. | Write a C# test method that exports the same workbook to HTML with ExportGridLines set to false and confirms that the HTML does not contain any border CSS. | Provide code that uses a MemoryStream to capture the HTML output from Aspose.Cells and checks for the presence or absence of the word "border" based on the ExportGridLines flag.
// Common Searches: aspocells html export unit test for gridlines in C# | how to assert border CSS in HTML output from Aspose.Cells | C# test ExportGridLines option produces gridline borders | verify Aspose.Cells HtmlSaveOptions ExportGridLines behavior with memory stream | unit testing Aspose.Cells HTML export without visual inspection
// Tags: Aspose.Cells HTML export gridlines unit test | HtmlSaveOptions ExportGridLines verification | C# memory stream HTML output testing | border CSS detection Aspose.Cells | gridline styling validation in HTML

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates a workbook, exports it to HTML twice—once with HtmlSaveOptions.ExportGridLines true and once false—captures the HTML via MemoryStream, and asserts that the HTML contains the word "border" only when gridlines are enabled, demonstrating how to unit‑test gridline rendering in Aspose.Cells.
    public class HtmlExportGridLinesTests
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and add some data
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Sample");
                worksheet.Cells["B2"].PutValue(12345);

                // Export to HTML with ExportGridLines = true
                var optionsWithGridLines = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportGridLines = true
                };
                string htmlWithGridLines;
                using (var streamWithGridLines = new MemoryStream())
                {
                    workbook.Save(streamWithGridLines, optionsWithGridLines);
                    htmlWithGridLines = Encoding.UTF8.GetString(streamWithGridLines.ToArray());
                }

                // Export to HTML with ExportGridLines = false
                var optionsWithoutGridLines = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportGridLines = false
                };
                string htmlWithoutGridLines;
                using (var streamWithoutGridLines = new MemoryStream())
                {
                    workbook.Save(streamWithoutGridLines, optionsWithoutGridLines);
                    htmlWithoutGridLines = Encoding.UTF8.GetString(streamWithoutGridLines.ToArray());
                }

                // Verify presence or absence of border styling
                bool containsBorder = htmlWithGridLines.Contains("border", StringComparison.OrdinalIgnoreCase);
                bool containsBorderWithout = htmlWithoutGridLines.Contains("border", StringComparison.OrdinalIgnoreCase);

                Console.WriteLine(containsBorder
                    ? "PASS: HTML with grid lines contains border style."
                    : "FAIL: HTML with grid lines should contain border style.");

                Console.WriteLine(!containsBorderWithout
                    ? "PASS: HTML without grid lines does not contain border style."
                    : "FAIL: HTML without grid lines should not contain border style.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
