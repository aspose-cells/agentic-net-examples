using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options for MHTML with IE compatibility
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);
        saveOptions.IsIECompatible = true; // Preserve formatting for IE

        // Save the workbook as an IE‑compatible MHTML file
        workbook.Save("output.mht", saveOptions);
    }
}