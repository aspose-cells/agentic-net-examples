// Title: C# unit test to confirm Aspose.Cells HTML export adds gridline borders only when ExportGridLines is enabled
// Common Searches: how to write a C# unit test for Aspose.Cells ExportGridLines HTML option | verify that Aspose.Cells HTML export includes border tags when ExportGridLines is true | Aspose.Cells HTMLSaveOptions ExportGridLines false excludes gridlines in generated HTML | unit testing Aspose.Cells HTML output for gridline rendering using in-memory stream | C# check for border attribute in Aspose.Cells exported HTML
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines testing | C# unit test Aspose.Cells HTML export | gridline border verification Aspose.Cells | in-memory stream provider Aspose.Cells HTML | validate HTML border markup from workbook export

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Simple test runner without external test frameworks
    // Provides a self‑contained C# example that creates a sample workbook, exports it to HTML with HtmlSaveOptions toggling the ExportGridLines flag, and uses simple assertions to ensure the generated HTML contains border markup when the flag is true and omits it when false, employing an in‑memory stream provider for stream‑based export.
    internal class Program
    {
        static void Main()
        {
            try
            {
                var tester = new HtmlExportGridLinesTests();

                tester.ExportGridLines_True_IncludesGridlines();
                Console.WriteLine("ExportGridLines_True_IncludesGridlines passed.");

                tester.ExportGridLines_False_ExcludesGridlines();
                Console.WriteLine("ExportGridLines_False_ExcludesGridlines passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }
    }

    public class HtmlExportGridLinesTests
    {
        // Helper method to create a simple workbook with some data
        private Workbook CreateSampleWorkbook()
        {
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate a 3x3 range with sample values
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            return workbook;
        }

        // Helper method to export a workbook to HTML string with specified ExportGridLines option
        private string ExportToHtml(Workbook workbook, bool exportGridLines)
        {
            var options = new HtmlSaveOptions
            {
                ExportGridLines = exportGridLines,
                ExportImagesAsBase64 = true,
                StreamProvider = new MemoryStreamProvider()
            };

            try
            {
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, options);
                    ms.Position = 0;
                    using (var reader = new StreamReader(ms))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to export workbook to HTML.", ex);
            }
        }

        // Test that when ExportGridLines is true, the generated HTML contains gridline styling (border attributes)
        public void ExportGridLines_True_IncludesGridlines()
        {
            var workbook = CreateSampleWorkbook();

            string html = ExportToHtml(workbook, exportGridLines: true);

            // Aspose.Cells adds border styles to <td> elements when gridlines are exported.
            if (!html.Contains("border"))
                throw new InvalidOperationException("HTML should contain gridline (border) definitions when ExportGridLines is true.");
        }

        // Test that when ExportGridLines is false, the generated HTML does not contain gridline styling
        public void ExportGridLines_False_ExcludesGridlines()
        {
            var workbook = CreateSampleWorkbook();

            string html = ExportToHtml(workbook, exportGridLines: false);

            // When gridlines are not exported, the HTML should not contain any "border" definitions.
            if (html.Contains("border"))
                throw new InvalidOperationException("HTML should not contain gridline (border) definitions when ExportGridLines is false.");
        }
    }

    // Simple in‑memory stream provider required by Aspose.Cells when exporting to a stream
    internal class MemoryStreamProvider : IStreamProvider
    {
        // Called by Aspose.Cells to obtain a stream for a given resource (e.g., the main HTML file)
        public void InitStream(StreamProviderOptions options)
        {
            // Provide a new MemoryStream for the resource.
            options.Stream = new MemoryStream();
        }

        // Called by Aspose.Cells after writing is complete to allow cleanup
        public void CloseStream(StreamProviderOptions options)
        {
            options.Stream?.Dispose();
        }
    }
}
