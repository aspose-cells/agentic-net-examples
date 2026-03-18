using System;
using Aspose.Cells;

class ExportWorkbookToHtml
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Keep conditional formatting rules as they are (do not merge them)
        htmlOptions.MergeAreas = false;

        // Save the workbook as HTML
        string outputPath = "output.html";
        workbook.Save(outputPath, htmlOptions);
    }
}