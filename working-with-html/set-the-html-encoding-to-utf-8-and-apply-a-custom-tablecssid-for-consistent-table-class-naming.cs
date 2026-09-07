// Title: Export an Aspose.Cells workbook to HTML with UTF-8 encoding and a custom TableCssId in C#
// AI Prompts: Write C# that saves a Workbook as an HTML file, setting HtmlSaveOptions.Encoding to UTF-8 and HtmlSaveOptions.TableCssId to a user-defined value. | Show how to configure Aspose.Cells HtmlSaveOptions to apply a specific CSS ID to the generated HTML table while ensuring the output uses UTF-8 character encoding.
// Common Searches: c# aspose.cells export excel to html with utf-8 encoding and custom table css id | how to set TableCssId in HtmlSaveOptions when saving workbook as html using Aspose.Cells | Aspose.Cells HtmlSaveOptions example for UTF8 output and table identifier in C#
// Tags: Aspose.Cells HtmlSaveOptions UTF-8 encoding | Aspose.Cells custom TableCssId for HTML export | C# export workbook to HTML with specific CSS identifier | HTML table styling using Aspose.Cells HtmlSaveOptions | set HTML file encoding Aspose.Cells C#

using System.Text;
using Aspose.Cells;

// Creates a workbook, adds sample data, configures HtmlSaveOptions with UTF-8 encoding and a custom TableCssId, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Example data – you can replace this with your own content
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Text");

        // Configure HTML export options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Set the HTML file encoding to UTF-8
            Encoding = Encoding.UTF8,

            // Apply a custom CSS ID to the generated HTML table for consistent styling
            TableCssId = "customTableId"
        };

        // Save the workbook as an HTML file using the specified options
        workbook.Save("ExportedDocument.html", htmlOptions);
    }
}
