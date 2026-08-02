using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class IconSetHtmlExport
{
    static void Main()
    {
        // Load the workbook that already contains IconSet conditional formatting rules
        string inputPath = "input.xlsx";               // path to the source workbook
        Workbook workbook = new Workbook(inputPath);    // create and load

        // Create HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Explicitly ensure that all data, including conditional formatting icons, are exported
        // (HtmlExportDataOptions.All is the default, but we set it to satisfy the requirement)
        htmlOptions.ExportDataOptions = HtmlExportDataOptions.All;

        // Save the workbook as HTML – icons from the IconSet will be rendered in the output
        string outputPath = "output.html";
        workbook.Save(outputPath, htmlOptions);

        Console.WriteLine($"Workbook saved as HTML with IconSet icons at: {outputPath}");
    }
}