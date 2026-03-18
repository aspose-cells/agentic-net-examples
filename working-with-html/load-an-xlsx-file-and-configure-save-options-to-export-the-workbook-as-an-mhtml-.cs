using System;
using Aspose.Cells;

namespace AsposeCellsMhtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options specifying MHTML format
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);
            // Optional: embed images as Base64 to keep them inside the MHTML file
            saveOptions.ExportImagesAsBase64 = true;

            // Save the workbook as an MHTML document
            workbook.Save("output.mht", saveOptions);
        }
    }
}