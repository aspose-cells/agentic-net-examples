// Title: Export a workbook to HTML with inline CSS definitions prefixed by a custom TableCssId using Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets HtmlSaveOptions.IsInlineCss = true and assigns a TableCssId value so that all table styles are emitted as inline CSS with the specified prefix. | Modify the sample workbook export to enable inline CSS, set TableCssId to a custom identifier, and ensure the resulting HTML contains the stylesheet directly for immediate preview.
// Common Searches: Aspose.Cells C# export HTML with inline CSS and custom TableCssId | How to enable IsInlineCss in HtmlSaveOptions for Aspose.Cells | Add a CSS class prefix to tables when saving workbook as HTML using Aspose.Cells | Preview exported HTML with embedded styles in .NET Aspose.Cells | C# Aspose.Cells HTML export with embedded stylesheet
// Tags: Aspose.Cells HtmlSaveOptions inline CSS | TableCssId custom class prefix Aspose.Cells | C# export workbook to HTML with embedded stylesheet | preview HTML with inline styles Aspose.Cells | set IsInlineCss property Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills it with sample data, configures HtmlSaveOptions to embed CSS inline and apply a custom TableCssId prefix, and saves the workbook as an HTML file for immediate preview.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.20);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);

            // Configure HTML export options (using default settings)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Define the output file path
            string outputPath = "ExportedWithInlineCss.html";

            // Save the workbook to HTML using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
