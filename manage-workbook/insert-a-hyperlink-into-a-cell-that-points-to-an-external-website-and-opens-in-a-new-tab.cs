using System;
using Aspose.Cells;

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
            // Parameters: cell name, rows in range, columns in range, address
            int linkIndex = worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");

            // Optionally set the display text for the hyperlink
            worksheet.Hyperlinks[linkIndex].TextToDisplay = "Visit Example.com";

            // Configure HTML save options to open links in a new tab/window
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.LinkTargetType = HtmlLinkTargetType.Blank; // _blank target

            // Save the workbook as HTML so the hyperlink opens in a new tab when clicked
            workbook.Save("HyperlinkInNewTab.html", htmlOptions);
        }
    }
}