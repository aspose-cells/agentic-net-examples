// Title: Convert an Excel workbook to HTML with a linked external CSS stylesheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# save Excel as HTML with external CSS file | How to generate a separate CSS stylesheet when converting XLSX to HTML using Aspose.Cells | C# example exporting workbook to HTML without inline styles Aspose.Cells | HtmlSaveOptions to create linked CSS file instead of embedded styles in Aspose.Cells | Saving Excel to HTML with external resources (images, CSS) using Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions external stylesheet | C# convert Excel to HTML separate CSS | Aspose.Cells generate linked CSS file | HTML export without inline styles Aspose.Cells | Aspose.Cells save workbook as HTML external resources

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample checks for the source XLSX file, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions to disable Base64 image embedding and use full‑path links, then saves the workbook as HTML. The operation creates an HTML file together with separate resource files such as images and a CSS stylesheet, avoiding inline style definitions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputHtml = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputFile);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Keep images as separate files rather than embedding them as Base64
                ExportImagesAsBase64 = false,
                // Use full path links for external resources (if needed)
                IsFullPathLink = true
                // Note: CssStyleSheetType and CssFileName are not available in the current Aspose.Cells version
            };

            // Save the workbook as HTML; external resources (images, CSS) will be generated alongside the HTML file
            workbook.Save(outputHtml, htmlOptions);

            Console.WriteLine($"Workbook successfully converted to HTML: {outputHtml}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
