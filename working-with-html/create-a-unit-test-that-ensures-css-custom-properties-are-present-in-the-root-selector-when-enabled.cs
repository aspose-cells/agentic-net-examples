// Title: Create a C# unit test with Aspose.Cells to verify background‑color CSS is emitted in the HTML export
// Common Searches: how to write a unit test for Aspose.Cells HTML export in C# | verify generated HTML contains background-color style using Aspose.Cells | C# Aspose.Cells test for CSS output of cell formatting | unit testing HtmlSaveOptions to ensure style rendering in .NET | Aspose.Cells example checking CSS properties in saved HTML
// Tags: Aspose.Cells HTML export unit test | verify CSS style generation Aspose.Cells | C# workbook to HTML background color validation | HtmlSaveOptions style rendering test | cell formatting to CSS verification Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsTests
{
    // The example creates a Workbook, applies a yellow background style to cell A1, saves the workbook to HTML with HtmlSaveOptions, reads the HTML from a memory stream, and asserts that the expected background‑color declaration is present, throwing an exception if it is missing.
    public class CssCustomPropertiesDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and add a style to a cell
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                var cell = sheet.Cells["A1"];
                cell.PutValue("Test");

                // Apply a background color style to generate a CSS custom property (if supported)
                var style = cell.GetStyle();
                style.ForegroundColor = System.Drawing.Color.Yellow;
                style.Pattern = BackgroundType.Solid;
                cell.SetStyle(style);

                // Configure HTML save options
                var saveOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportActiveWorksheetOnly = true
                    // Note: ExportCssCustomProperties is not available in the current Aspose.Cells version.
                };

                // Save the workbook to a memory stream
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, saveOptions);
                    stream.Position = 0;

                    // Read the generated HTML content
                    string htmlContent;
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        htmlContent = reader.ReadToEnd();
                    }

                    // Verify that the generated HTML contains the expected background color style
                    if (!htmlContent.Contains("background-color: yellow", StringComparison.OrdinalIgnoreCase) &&
                        !htmlContent.Contains("background-color:#ffff00", StringComparison.OrdinalIgnoreCase))
                    {
                        throw new Exception("The expected background color style was not found in the generated HTML.");
                    }

                    Console.WriteLine("Test passed. Generated HTML contains the expected background color style.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
