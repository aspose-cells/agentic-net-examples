using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to enable the CrossHideRight behavior
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.HtmlCrossStringType = HtmlCrossType.CrossHideRight; // hide right string when texts overlap

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}