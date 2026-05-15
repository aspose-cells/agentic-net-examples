using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsHyperlinkExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some visible text to the cell where the hyperlink will be placed
            worksheet.Cells["A1"].PutValue("Visit Aspose");

            // Insert a hyperlink that points to an external website (e.g., https://www.aspose.com)
            // The Add method parameters: cell name, total rows, total columns, address
            int hyperlinkIndex = worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

            // Optionally customize the displayed text (default is the address)
            worksheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Aspose Website";

            // Optionally add a screen tip
            worksheet.Hyperlinks[hyperlinkIndex].ScreenTip = "Open Aspose in a new tab";

            // When saving to HTML, set the link target type to Blank so the link opens in a new window/tab
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.LinkTargetType = HtmlLinkTargetType.Blank; // Opens in new tab/window

            // Save the workbook as HTML (the hyperlink will open in a new tab when clicked)
            workbook.Save("HyperlinkInNewTab.html", saveOptions);
        }
    }
}