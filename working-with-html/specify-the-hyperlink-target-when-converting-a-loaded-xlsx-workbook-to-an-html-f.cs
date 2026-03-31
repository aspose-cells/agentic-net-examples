using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Add a sample hyperlink (optional, to see the target effect)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Visit Aspose");
        sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

        // Create HTML save options and set the hyperlink target type
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.LinkTargetType = HtmlLinkTargetType.Blank; // Opens links in a new window/tab

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}