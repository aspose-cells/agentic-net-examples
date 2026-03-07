using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsRtlExport
{
    class Program
    {
        static void Main()
        {
            // Load the source XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);
        }
    }
}