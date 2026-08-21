// Title: C# – Set a Custom HTML Page Title with Aspose.Cells HtmlSaveOptions (.NET)
// Description: Learn how to export an Excel workbook to HTML and define a custom <title> tag using Aspose.Cells HtmlSaveOptions.PageTitle in C#. The example creates a workbook, writes data, sets the page title, and saves the file as custom_title.html.
// Keywords: Aspose.Cells HtmlSaveOptions | C# set HTML page title | Aspose.Cells export to HTML | PageTitle property .NET | Excel to HTML conversion | custom HTML title Aspose | Aspose.Cells sample code | save workbook as HTML C# | HTML title tag Excel export | Aspose.Cells tutorial
// Common Searches: Aspose.Cells set HTML title C# | HtmlSaveOptions PageTitle example | export Excel to HTML with custom title | C# Aspose.Cells HTML export settings | how to change <title> tag when saving as HTML
// Developer Intent: Apply the HtmlSaveOptions.PageTitle property to assign a specific <title> value to the HTML file generated from an Excel workbook using Aspose.Cells in .NET.
// Use Cases: Create SEO‑friendly web pages from Excel reports by giving each HTML file a descriptive title. | Automate batch conversion of workbooks to HTML where the title reflects the worksheet name or report period. | Generate printable HTML dashboards with meaningful titles for easier bookmarking and sharing.
// AI Prompts: Provide C# code that converts an Excel workbook to HTML and sets a custom page title with Aspose.Cells. | Show how to use HtmlSaveOptions to set the HTML <title> and attach an external CSS file during export. | Explain how to loop through multiple workbooks and assign a unique PageTitle for each HTML output.

using System;
using Aspose.Cells;

// Learn how to export an Excel workbook to HTML and define a custom <title> tag using Aspose.Cells HtmlSaveOptions.PageTitle in C#. The example creates a workbook, writes data, sets the page title, and saves the file as custom_title.html.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello Aspose.Cells!");

        // Create HTML save options (uses HtmlSaveOptions constructor)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Set a custom page title for the generated HTML
        saveOptions.PageTitle = "My Custom HTML Title";

        // Save the workbook as an HTML file with the specified title
        string outputPath = "custom_title.html";
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("HTML file saved to: " + outputPath);
    }
}
