using System;
using Aspose.Cells;

namespace AsposeCellsHtmlToPdf
{
    class Program
    {
        static void Main()
        {
            // Create HtmlLoadOptions instance
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Optional: enable support for <div> tags in the HTML
            loadOptions.SupportDivTag = true;

            // Load the HTML file into a Workbook using the load options
            Workbook workbook = new Workbook("input.html", loadOptions);

            // Save the loaded workbook as a PDF file
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}