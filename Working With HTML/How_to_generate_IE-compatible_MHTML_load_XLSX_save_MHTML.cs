using System;
using Aspose.Cells;

class GenerateMhtml
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options specifying MHTML format
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);
        // Enable IE compatibility for the generated MHTML
        saveOptions.IsIECompatible = true;

        // Save the workbook as an IE‑compatible MHTML file
        workbook.Save("output.mht", saveOptions);
    }
}