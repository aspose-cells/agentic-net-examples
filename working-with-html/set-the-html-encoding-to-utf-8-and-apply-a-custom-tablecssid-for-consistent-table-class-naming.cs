// Title: Save an Excel workbook as HTML with UTF-8 encoding and a custom TableCssId using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and exports it to HTML with UTF-8 encoding and a user-defined TableCssId using Aspose.Cells. | Show how to set HtmlSaveOptions.Encoding and HtmlSaveOptions.TableCssId before calling Workbook.Save in a .NET application.
// Common Searches: Aspose.Cells C# set HTML export character set to UTF-8 | How to assign a custom CSS ID to the generated HTML table with HtmlSaveOptions | Export Excel to HTML with specific table identifier using Aspose.Cells .NET | HtmlSaveOptions Encoding property example for UTF-8 in Aspose.Cells | Define TableCssId when saving workbook as HTML in C#
// Tags: Aspose.Cells HtmlSaveOptions UTF-8 encoding | Aspose.Cells HtmlSaveOptions custom TableCssId | C# export Excel to HTML Aspose.Cells | HTML table CSS identifier Aspose.Cells | UTF-8 character set Aspose.Cells HTML export

using System.Text;
using Aspose.Cells;

// Loads an Excel workbook, configures HtmlSaveOptions to use UTF-8 encoding and a custom TableCssId, then saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Set the HTML encoding to UTF-8
            Encoding = Encoding.UTF8,

            // Apply a custom CSS ID for the generated HTML table
            TableCssId = "customTableId"
        };

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
