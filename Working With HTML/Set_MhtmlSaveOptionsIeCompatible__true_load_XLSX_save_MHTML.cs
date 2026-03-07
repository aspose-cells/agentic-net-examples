using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create save options for MHTML format
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.MHtml);

        // Enable IE compatibility for the generated MHTML file
        mhtmlOptions.IsIECompatible = true;

        // Save the workbook as MHTML using the configured options
        workbook.Save("output.mht", mhtmlOptions);
    }
}