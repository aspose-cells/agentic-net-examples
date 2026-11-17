// Title: How to set UTF-8 encoding and a custom TableCssId when saving a workbook to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set HTML encoding to UTF-8 using HtmlSaveOptions | assign a custom TableCssId for HTML tables in Aspose.Cells .NET | save Excel workbook as HTML with a specific CSS id using Aspose.Cells | configure HtmlSaveOptions for UTF-8 encoding and table CSS identifier in C#
// Tags: HtmlSaveOptions UTF-8 | TableCssId Aspose.Cells | Excel workbook HTML conversion C# | Aspose.Cells HTML export options | custom HTML table identifier

using Aspose.Cells;
using System.Text;

// The example creates a workbook, adds sample data, configures HtmlSaveOptions to use UTF-8 encoding and sets TableCssId to "customTableCssId", then saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Example: add some data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Header 1");
        sheet.Cells["B1"].PutValue("Header 2");
        sheet.Cells["A2"].PutValue("Value 1");
        sheet.Cells["B2"].PutValue("Value 2");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.Encoding = Encoding.UTF8;          // Set HTML encoding to UTF-8
        htmlOptions.TableCssId = "customTableCssId";   // Apply custom TableCssId for consistent table class naming

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
