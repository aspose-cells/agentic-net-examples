using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load the source XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options and set link target to open in a new tab/window
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.LinkTargetType = HtmlLinkTargetType.Blank; // corresponds to target="_blank"

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", saveOptions);
        }
    }
}