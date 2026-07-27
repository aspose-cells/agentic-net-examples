// Title: Aspose.Cells for .NET – Set Custom HTML Page Title via HtmlSaveOptions
// Description: Shows how to set the HtmlSaveOptions.PageTitle property before exporting a Workbook to HTML, so the generated <title> tag contains the desired text.
// Keywords: Aspose.Cells | .NET | HtmlSaveOptions | PageTitle | HTML export | C# Excel to HTML | set HTML title | web page metadata | custom page heading
// Common Searches: Aspose.Cells set HTML title C# | HtmlSaveOptions PageTitle example | export Excel to HTML with custom <title> | how to change HTML page title using Aspose.Cells | C# generate HTML from workbook with specific title
// Developer Intent: Apply a user‑defined string to the <title> element of the HTML file produced from an Excel workbook.
// Use Cases: Create a web‑ready sales report where the browser tab shows a meaningful heading. | Batch‑convert worksheets to individual HTML pages, each reflecting its sheet name in the title bar. | Generate documentation from Excel templates with SEO‑friendly page titles for better search visibility.
// AI Prompts: Write C# code that uses Aspose.Cells to export a workbook to HTML and assigns a variable value to HtmlSaveOptions.PageTitle. | Explain the relationship between HtmlSaveOptions.PageTitle and the <title> tag, and describe the default behavior when the property is left unset.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to set the HtmlSaveOptions.PageTitle property before exporting a Workbook to HTML, so the generated <title> tag contains the desired text.
    public class HtmlPageTitleDemo
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello Aspose.Cells!");

            // Create HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Set a custom title for the generated HTML page
            saveOptions.PageTitle = "My Custom HTML Title";

            // Define the output HTML file path
            string outputPath = "CustomTitle.html";

            // Save the workbook as HTML using the specified options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"HTML file saved to '{outputPath}' with page title '{saveOptions.PageTitle}'.");
        }
    }
}
