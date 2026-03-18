using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportActiveWorksheetOnly = true;          // Export only the active worksheet
        htmlOptions.ExportDataOptions = HtmlExportDataOptions.All; // Export all data (including formulas, comments, etc.)
        htmlOptions.HtmlVersion = HtmlVersion.Html5;           // Use HTML5 standard
        htmlOptions.ExportImagesAsBase64 = true;              // Embed images as Base64 strings
        htmlOptions.SaveAsSingleFile = true;                  // Produce a single HTML file
        htmlOptions.PageTitle = "Exported Workbook";          // Set the HTML page title

        // Save the workbook as HTML using the configured options
        string outputPath = "output.html";
        workbook.Save(outputPath, htmlOptions);
    }
}