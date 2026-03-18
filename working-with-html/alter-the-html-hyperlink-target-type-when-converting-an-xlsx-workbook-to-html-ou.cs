using System;
using Aspose.Cells;

namespace AsposeCellsHtmlLinkTargetDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some text to a cell
            sheet.Cells["A1"].PutValue("Visit Aspose");

            // Add a hyperlink to the cell (A1)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlink address
            sheet.Hyperlinks.Add(0, 0, 1, 1, "https://www.aspose.com");

            // Create HTML save options and set the hyperlink target type
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Open links in a new window/tab (_blank)
            saveOptions.LinkTargetType = HtmlLinkTargetType.Blank;

            // Save the workbook as HTML using the configured options
            workbook.Save("AsposeLinkTarget.html", saveOptions);

            Console.WriteLine("HTML file saved with hyperlink target set to '_blank'.");
        }
    }
}