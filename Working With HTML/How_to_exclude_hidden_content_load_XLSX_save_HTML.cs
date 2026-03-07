using System;
using Aspose.Cells;

namespace AsposeCellsExcludeHiddenContent
{
    class Program
    {
        static void Main()
        {
            // Load the source XLSX workbook
            // (Replace "input.xlsx" with the actual path to your Excel file)
            Workbook workbook = new Workbook("input.xlsx");

            // Configure HTML save options to exclude hidden content
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Do not export hidden worksheets
                ExportHiddenWorksheet = false,

                // Remove hidden rows from the generated HTML (instead of rendering them as hidden)
                HiddenRowDisplayType = HtmlHiddenRowDisplayType.Remove,

                // Remove hidden columns from the generated HTML (instead of rendering them as hidden)
                HiddenColDisplayType = HtmlHiddenColDisplayType.Remove
            };

            // Save the workbook as HTML using the configured options
            // (Replace "output.html" with the desired output path)
            workbook.Save("output.html", saveOptions);
        }
    }
}