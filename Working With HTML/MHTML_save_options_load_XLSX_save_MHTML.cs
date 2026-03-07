using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook from disk
        Workbook workbook = new Workbook("input.xlsx");

        // Create HtmlSaveOptions specifying the MHTML format
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);
        // Optional: improve visual presentation
        saveOptions.PresentationPreference = true;
        // Optional: embed images directly in the MHTML file
        saveOptions.ExportImagesAsBase64 = true;

        // Save the workbook as an MHTML file using the options
        workbook.Save("output.mht", saveOptions);
    }
}