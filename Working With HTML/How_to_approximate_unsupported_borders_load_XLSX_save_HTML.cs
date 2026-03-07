using System;
using Aspose.Cells;

class ApproximateUnsupportedBorders
{
    static void Main()
    {
        // Load the source XLSX file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and enable border approximation
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportSimilarBorderStyle = true // Approximate borders not supported by browsers
        };

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}