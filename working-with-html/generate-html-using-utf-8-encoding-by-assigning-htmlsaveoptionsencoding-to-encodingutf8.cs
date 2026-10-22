// Title: Save an Aspose.Cells workbook as UTF-8 encoded HTML in C#
// AI Prompts: Write C# code that creates a Workbook, configures HtmlSaveOptions.Encoding to UTF8, and saves the workbook as an HTML file. | Show how to use Aspose.Cells HtmlSaveOptions to generate UTF-8 HTML output from a spreadsheet in a .NET application.
// Common Searches: Aspose.Cells export workbook to HTML with UTF-8 encoding C# example | how to set HtmlSaveOptions.Encoding to UTF8 in Aspose.Cells .NET | C# generate UTF-8 HTML from Excel using Aspose.Cells HtmlSaveOptions | save Excel file as UTF-8 HTML using Aspose.Cells library
// Tags: Aspose.Cells HtmlSaveOptions UTF-8 encoding | export workbook to HTML C# Aspose.Cells | set HTML encoding Aspose.Cells .NET | save Excel as UTF-8 HTML using Aspose

using System;
using System.Text;
using Aspose.Cells;

// The sample creates a Workbook, adds sample data, configures HtmlSaveOptions with Encoding = Encoding.UTF8, and saves the workbook as a UTF-8 encoded HTML file (output.html) using Aspose.Cells in C#.
class HtmlExportExample
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // creates a blank workbook
        // Optionally, you can load an existing workbook:
        // Workbook workbook = new Workbook("input.xlsx");

        // Add some data to the first worksheet for demonstration
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Configure HTML save options with UTF-8 encoding
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.Encoding = Encoding.UTF8; // set UTF-8 encoding

        // Save the workbook as an HTML file using the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
