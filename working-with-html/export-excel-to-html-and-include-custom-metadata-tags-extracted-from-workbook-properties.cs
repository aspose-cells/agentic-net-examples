// Title: Convert an Excel workbook to HTML and inject custom document property meta tags using Aspose.Cells for .NET
// Common Searches: how to include custom workbook properties as meta tags when saving Excel as HTML with Aspose.Cells | Aspose.Cells C# export worksheet to HTML with embedded metadata | add custom document properties to HTML head during Excel to HTML conversion in .NET | save Excel as HTML with base64 images and custom meta tags using Aspose.Cells | programmatically insert meta tags into generated HTML from Aspose.Cells workbook
// Tags: Aspose.Cells HtmlSaveOptions for HTML export | C# embed custom properties into HTML head | export active worksheet to HTML with base64 images | convert Excel to HTML preserving document metadata | programmatic insertion of meta tags in Aspose.Cells output

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, extracts its custom document properties, builds corresponding <meta> tags, saves the workbook as HTML using HtmlSaveOptions (exporting only the active worksheet and embedding images as base64), inserts the generated meta tags into the <head> element of the HTML, and writes the final file to disk.
class ExcelToHtmlWithMetadata
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Path to the output HTML file
        string htmlPath = "output.html";

        // Load the workbook
        Workbook workbook = new Workbook(excelPath);

        // Retrieve custom document properties
        var customProps = workbook.CustomDocumentProperties;

        // Build meta tags string from custom properties
        StringBuilder metaTagsBuilder = new StringBuilder();
        foreach (var prop in customProps)
        {
            // Ensure the property has a name and a value
            if (!string.IsNullOrEmpty(prop.Name) && prop.Value != null)
            {
                // Escape double quotes in the value
                string escapedValue = prop.Value.ToString().Replace("\"", "&quot;");
                metaTagsBuilder.AppendLine(
                    $"<meta name=\"{prop.Name}\" content=\"{escapedValue}\" />");
            }
        }
        string metaTags = metaTagsBuilder.ToString();

        // Prepare HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Export only the first worksheet (adjust as needed)
            ExportActiveWorksheetOnly = true,
            // Do not embed images as separate files
            ExportImagesAsBase64 = true
        };

        // Save workbook to a memory stream as HTML
        using (MemoryStream htmlStream = new MemoryStream())
        {
            workbook.Save(htmlStream, saveOptions);
            htmlStream.Position = 0;

            // Read the generated HTML
            string htmlContent;
            using (StreamReader reader = new StreamReader(htmlStream))
            {
                htmlContent = reader.ReadToEnd();
            }

            // Insert meta tags into the <head> section
            const string headTag = "<head>";
            int headIndex = htmlContent.IndexOf(headTag, StringComparison.OrdinalIgnoreCase);
            if (headIndex != -1)
            {
                int insertPos = headIndex + headTag.Length;
                htmlContent = htmlContent.Insert(insertPos, Environment.NewLine + metaTags);
            }
            else
            {
                // If <head> not found, prepend meta tags at the beginning
                htmlContent = metaTags + Environment.NewLine + htmlContent;
            }

            // Write the final HTML to the output file
            File.WriteAllText(htmlPath, htmlContent, Encoding.UTF8);
        }

        Console.WriteLine("Excel exported to HTML with custom metadata tags.");
    }
}
