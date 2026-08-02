using System;
using Aspose.Cells;

namespace HyperlinkPdfDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a hyperlink to cell A1 (row 0, column 0)
            sheet.Hyperlinks.Add(0, 0, 1, 1, "https://www.aspose.com");

            // Set display text and screen tip for the hyperlink
            Hyperlink link = sheet.Hyperlinks[0];
            link.TextToDisplay = "Aspose";
            link.ScreenTip = "Visit Aspose website";

            // Save the workbook as PDF; hyperlinks are preserved and clickable
            workbook.Save("HyperlinkDemo.pdf");
        }
    }
}