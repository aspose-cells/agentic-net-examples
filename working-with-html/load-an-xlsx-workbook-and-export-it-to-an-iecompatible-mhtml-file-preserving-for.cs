using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options for MHTML output
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);
        // Ensure the generated file is compatible with Internet Explorer
        saveOptions.IsIECompatible = true;
        // Preserve formatting, images, and interactive elements
        saveOptions.ExportImagesAsBase64 = true;
        saveOptions.ExportFrameScriptsAndProperties = true;
        saveOptions.ExportWorksheetProperties = true;
        saveOptions.ExportWorkbookProperties = true;
        saveOptions.ExportDocumentProperties = true;
        saveOptions.ExportFormula = true;
        saveOptions.ExportHiddenWorksheet = true;
        saveOptions.ExportGridLines = true;

        // Save the workbook as an IE‑compatible MHTML file
        workbook.Save("output.mht", saveOptions);
    }
}