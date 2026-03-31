using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Initialize HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Note: The ExpandRtlText property is not available in this version of Aspose.Cells.
        // If RTL text expansion is required, ensure you are using a version that supports it.

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}