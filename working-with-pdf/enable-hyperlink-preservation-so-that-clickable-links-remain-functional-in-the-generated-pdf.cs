using System;
using Aspose.Cells;

namespace HyperlinkPdfDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put some display text into cell A1
            sheet.Cells["A1"].PutValue("Visit Aspose");

            // Add a hyperlink to cell A1 (row 0, column 0)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlink address
            sheet.Hyperlinks.Add(0, 0, 1, 1, "https://www.aspose.com");

            // Optionally, customize the hyperlink's display text and screen tip
            Hyperlink link = sheet.Hyperlinks[0];
            link.TextToDisplay = "Aspose";
            link.ScreenTip = "Click to open Aspose website";

            // Save the workbook as PDF; hyperlinks are preserved by default
            workbook.Save("HyperlinkPreserved.pdf", SaveFormat.Pdf);
        }
    }
}