using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath);

        // Set up HTML save options to keep formatting and related properties
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportFormula = true;               // keep formulas
        htmlOptions.ExportGridLines = true;             // show grid lines
        htmlOptions.ExportWorksheetProperties = true;   // preserve worksheet settings
        htmlOptions.ExportWorkbookProperties = true;    // preserve workbook metadata
        htmlOptions.HtmlVersion = HtmlVersion.Html5;    // use modern HTML5 output

        // Export the workbook to an HTML file
        string outputPath = "output.html";
        workbook.Save(outputPath, htmlOptions);
    }
}