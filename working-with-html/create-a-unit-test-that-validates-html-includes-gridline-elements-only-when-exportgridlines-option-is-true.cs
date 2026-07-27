// Title: C# Unit Test for Aspose.Cells HTML Export – Verify Gridlines Appear Only When ExportGridLines Is True
// Description: Creates a workbook, sets worksheet gridlines visible, saves two HTML files with HtmlSaveOptions (ExportGridLines = true/false), reads the outputs, and asserts that gridline‑related CSS or markup exists only when the option is true. The test cleans up temporary files after execution.
// Keywords: Aspose.Cells | HTML export | ExportGridLines | C# unit test | gridline validation | HtmlSaveOptions | automated testing | CI validation | border CSS check
// Common Searches: Aspose.Cells unit test for ExportGridLines | how to verify gridlines in exported HTML using C# | NUnit test Aspose.Cells HTML gridline option | check border CSS in Aspose.Cells HTML output | validate ExportGridLines false removes gridlines
// Developer Intent: Write an automated test that confirms gridline markup is generated only when HtmlSaveOptions.ExportGridLines is set to true.
// Use Cases: Regression testing to guarantee visual consistency of HTML reports. | CI pipeline check that prevents unwanted borders in exported HTML. | Documentation example showing how to validate Aspose.Cells export settings.
// AI Prompts: Generate an NUnit test method that creates a workbook, exports HTML with ExportGridLines true and false, reads the files, asserts presence or absence of gridline CSS, and deletes the temporary files. | Write a MSTest case for Aspose.Cells that verifies the ExportGridLines flag influences HTML output by searching for 'gridline' or 'border' strings. | Provide an xUnit example that checks gridline markup in generated HTML based on ExportGridLines setting and includes proper file cleanup.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Creates a workbook, sets worksheet gridlines visible, saves two HTML files with HtmlSaveOptions (ExportGridLines = true/false), reads the outputs, and asserts that gridline‑related CSS or markup exists only when the option is true. The test cleans up temporary files after execution.
    class Program
    {
        private const string OutputWithGridLines = "gridlines_true.html";
        private const string OutputWithoutGridLines = "gridlines_false.html";

        static void Main()
        {
            try
            {
                // Create a workbook and add some data
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.IsGridlinesVisible = true; // make gridlines visible in the worksheet
                worksheet.Cells["A1"].PutValue("Sample");
                worksheet.Cells["B2"].PutValue(123);

                // Save HTML with ExportGridLines = true
                HtmlSaveOptions optionsWithGridLines = new HtmlSaveOptions
                {
                    ExportGridLines = true,
                    ExportActiveWorksheetOnly = true
                };
                workbook.Save(OutputWithGridLines, optionsWithGridLines);

                // Save HTML with ExportGridLines = false
                HtmlSaveOptions optionsWithoutGridLines = new HtmlSaveOptions
                {
                    ExportGridLines = false,
                    ExportActiveWorksheetOnly = true
                };
                workbook.Save(OutputWithoutGridLines, optionsWithoutGridLines);

                // Ensure files were created
                if (!File.Exists(OutputWithGridLines) || !File.Exists(OutputWithoutGridLines))
                    throw new FileNotFoundException("One or both HTML output files were not created.");

                // Read the generated HTML files
                string htmlWithGridLines = File.ReadAllText(OutputWithGridLines);
                string htmlWithoutGridLines = File.ReadAllText(OutputWithoutGridLines);

                // Verify that the HTML containing gridlines has gridline‑related CSS or markup
                bool hasGridLines = htmlWithGridLines.Contains("gridline", StringComparison.OrdinalIgnoreCase) ||
                                    htmlWithGridLines.Contains("border", StringComparison.OrdinalIgnoreCase);
                if (!hasGridLines)
                    throw new InvalidOperationException("HTML should contain gridline definitions when ExportGridLines is true.");

                // Verify that the HTML without gridlines does NOT contain those definitions
                bool hasGridLinesInFalse = htmlWithoutGridLines.Contains("gridline", StringComparison.OrdinalIgnoreCase) ||
                                           htmlWithoutGridLines.Contains("border", StringComparison.OrdinalIgnoreCase);
                if (hasGridLinesInFalse)
                    throw new InvalidOperationException("HTML should not contain gridline definitions when ExportGridLines is false.");

                Console.WriteLine("Gridline export test passed.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Clean up generated files
                try { if (File.Exists(OutputWithGridLines)) File.Delete(OutputWithGridLines); } catch { }
                try { if (File.Exists(OutputWithoutGridLines)) File.Delete(OutputWithoutGridLines); } catch { }
            }
        }
    }
}
