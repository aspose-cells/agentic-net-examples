using System;
using Aspose.Cells;

namespace AsposeCellsLinkTargetDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // (Optional) Add a hyperlink to demonstrate the target attribute
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Visit Aspose");
            sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

            // Configure HTML save options to set the link target type
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Open links in a new window/tab (_blank)
            htmlOptions.LinkTargetType = HtmlLinkTargetType.Blank;

            // Save the workbook as HTML with the specified link target
            workbook.Save("output.html", htmlOptions);
        }
    }
}