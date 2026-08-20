// Title: Embed Worksheet CSS in HTML using Aspose.Cells HtmlSaveOptions (ExportWorksheetCSSSeparately = false)
// Description: Shows how to create a workbook, set HtmlSaveOptions.ExportWorksheetCSSSeparately to false so the stylesheet is placed inside a <style> tag, save the workbook as HTML, and programmatically verify that the CSS is embedded.
// Keywords: Aspose.Cells HTML export | ExportWorksheetCSSSeparately | embed CSS in HTML | inline stylesheet Aspose.Cells | C# Aspose.Cells HTML save | verify embedded CSS | no external CSS file | HTMLSaveOptions CSS embedding
// Common Searches: Aspose.Cells embed CSS in exported HTML | HtmlSaveOptions ExportWorksheetCSSSeparately example | How to include stylesheet inside HTML with Aspose.Cells | Check for <style> tag in Aspose.Cells HTML output | Save Excel as HTML with inline CSS C#
// Developer Intent: Generate an HTML file from a workbook where all CSS is embedded directly in the document, eliminating external stylesheet files.
// Use Cases: Create a single, portable HTML report that contains all styling. | Produce HTML email content where external CSS links are prohibited. | Automate validation that the exported HTML includes a <style> block.
// AI Prompts: Modify the example to embed images as base64 strings while keeping CSS inline. | Add code that asserts a specific CSS rule exists inside the generated <style> tag. | Explain how to combine CSS embedding with custom page layout and image handling options in HtmlSaveOptions.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExportDemo
{
    // Shows how to create a workbook, set HtmlSaveOptions.ExportWorksheetCSSSeparately to false so the stylesheet is placed inside a <style> tag, save the workbook as HTML, and programmatically verify that the CSS is embedded.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Embedded CSS Test");

            // Initialize HtmlSaveOptions and ensure CSS is embedded (ExportWorksheetCSSSeparately = false)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportWorksheetCSSSeparately = false; // CSS will be embedded in the HTML file

            // Define output HTML file path
            string outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "EmbeddedCssOutput.html");

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, saveOptions);

            // Verify that CSS is embedded by checking for a <style> tag in the generated HTML
            string htmlContent = File.ReadAllText(outputPath);
            bool cssEmbedded = htmlContent.Contains("<style", StringComparison.OrdinalIgnoreCase);

            Console.WriteLine($"HTML file saved to: {outputPath}");
            Console.WriteLine($"CSS embedded within HTML: {cssEmbedded}");
        }
    }
}
