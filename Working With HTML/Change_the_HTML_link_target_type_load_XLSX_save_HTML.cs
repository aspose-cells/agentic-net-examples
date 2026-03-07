using System;
using Aspose.Cells;

namespace AsposeCellsHtmlLinkTargetDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing XLSX workbook
            // Replace "input.xlsx" with the path to your source Excel file
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Set the target type for hyperlinks in the generated HTML.
            // Options: HtmlLinkTargetType.Blank, Parent, Self, Top
            // Here we set it to open links in a new window/tab.
            htmlOptions.LinkTargetType = HtmlLinkTargetType.Blank;

            // Save the workbook as an HTML file using the configured options
            // Replace "output.html" with the desired output path
            workbook.Save("output.html", htmlOptions);
        }
    }
}