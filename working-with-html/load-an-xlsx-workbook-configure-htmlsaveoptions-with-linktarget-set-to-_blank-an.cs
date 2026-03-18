using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Set the link target type to "_blank" (open links in a new window/tab)
        saveOptions.LinkTargetType = HtmlLinkTargetType.Blank;

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", saveOptions);
    }
}