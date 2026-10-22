// Title: Write a C# unit test that confirms Aspose.Cells HTML export adds gridline CSS only when ExportGridLines is true
// AI Prompts: Generate an MSTest or NUnit test method that saves a workbook to HTML with HtmlSaveOptions.ExportGridLines set to true and asserts that the output contains a "border:" style. | Create a second test case that saves the same workbook with ExportGridLines set to false and asserts that the generated HTML does not contain any "border:" CSS.
// Common Searches: c# unit test Aspose.Cells HtmlSaveOptions ExportGridLines true false | how to assert border CSS in HTML produced by Aspose.Cells | testing gridline rendering in Aspose.Cells HTML export | verify Aspose.Cells HTML output includes gridline styles only when enabled | write NUnit test for Aspose.Cells HTML export gridlines option
// Tags: Aspose.Cells HTML gridline unit test | HtmlSaveOptions ExportGridLines verification C# | C# test for Aspose.Cells HTML border CSS | validate gridline rendering in HTML export | unit testing Aspose.Cells HTML output

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a C# unit test that exports a workbook to HTML with HtmlSaveOptions.ExportGridLines toggled, then checks the resulting HTML for the presence or absence of "border:" CSS to ensure gridlines are rendered only when the option is enabled.
    class HtmlExportGridLinesDemo
    {
        // Creates a simple workbook with a few cells populated.
        private static Workbook CreateSampleWorkbook()
        {
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);
            return workbook;
        }

        // Exports the given workbook to an HTML string using the specified ExportGridLines setting.
        private static string ExportToHtml(Workbook workbook, bool exportGridLines)
        {
            var options = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportGridLines = exportGridLines
            };

            using (var stream = new MemoryStream())
            {
                workbook.Save(stream, options);
                stream.Position = 0;
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        // Checks whether the generated HTML contains border styles when ExportGridLines is true.
        private static void VerifyExportGridLinesTrue()
        {
            var workbook = CreateSampleWorkbook();
            string html = ExportToHtml(workbook, true);
            if (html.Contains("border:"))
                Console.WriteLine("PASS: HTML contains border styles when ExportGridLines is true.");
            else
                Console.WriteLine("FAIL: HTML does NOT contain border styles when ExportGridLines is true.");
        }

        // Checks whether the generated HTML does NOT contain border styles when ExportGridLines is false.
        private static void VerifyExportGridLinesFalse()
        {
            var workbook = CreateSampleWorkbook();
            string html = ExportToHtml(workbook, false);
            if (!html.Contains("border:"))
                Console.WriteLine("PASS: HTML does NOT contain border styles when ExportGridLines is false.");
            else
                Console.WriteLine("FAIL: HTML contains border styles when ExportGridLines is false.");
        }

        static void Main(string[] args)
        {
            try
            {
                VerifyExportGridLinesTrue();
                VerifyExportGridLinesFalse();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
