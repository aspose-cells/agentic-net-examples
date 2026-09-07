// Title: Embed CSS in HTML output when saving an Aspose.Cells workbook by disabling ExportWorksheetCSSSeparately (C#)
// AI Prompts: Generate C# code that saves a Workbook to an HTML file with inline CSS by setting HtmlSaveOptions.ExportWorksheetCSSSeparately = false. | Write C# to load the saved HTML file and programmatically verify that a <style> element exists, confirming CSS is embedded. | Adapt the example to write the HTML output to a MemoryStream while keeping the CSS inline, and return the HTML string.
// Common Searches: Aspose.Cells C# save workbook as HTML with embedded CSS | How to prevent external CSS file when exporting Excel to HTML using Aspose.Cells | HtmlSaveOptions ExportWorksheetCSSSeparately false example in C# | Check for inline style tag in Aspose.Cells generated HTML | Save Excel to HTML without separate stylesheet using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions inline stylesheet | embed CSS in Aspose.Cells HTML export | C# save workbook to HTML with embedded stylesheet | verify inline style tag in generated HTML | memory stream HTML export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates creating a workbook, adding data, configuring HtmlSaveOptions with ExportWorksheetCSSSeparately = false to embed CSS directly in the generated HTML, saving the file, and programmatically confirming the presence of a <style> tag.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and add some data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "SampleSheet";
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(1.25);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(0.75);

        // Configure HTML save options to embed CSS within the HTML
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportWorksheetCSSSeparately = false; // CSS will be embedded

        // Define output HTML file path
        string htmlFile = "SampleOutput.html";

        // Save the workbook as HTML using the configured options
        workbook.Save(htmlFile, htmlOptions);

        // Verify that CSS is embedded by checking for a <style> tag in the generated HTML
        string htmlContent = File.ReadAllText(htmlFile);
        bool isCssEmbedded = htmlContent.Contains("<style");

        Console.WriteLine($"CSS embedded in HTML: {isCssEmbedded}");
    }
}
