using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace HyperlinkExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a hyperlink to cell A1 that points to an external website
            // Parameters: cell name, rows, columns, address
            int linkIndex = worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");

            // Set the display text for the hyperlink (optional)
            worksheet.Hyperlinks[linkIndex].TextToDisplay = "Visit Example.com";

            // When saving to HTML, set the link target type to open in a new tab/window
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.LinkTargetType = HtmlLinkTargetType.Blank; // "_blank" target

            // Save the workbook as an HTML file; the hyperlink will open in a new tab
            workbook.Save("HyperlinkExample.html", saveOptions);
        }
    }
}