using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        string sourceFile = "input.xlsx";
        Workbook workbook = new Workbook(sourceFile);

        // Create HTML save options and enable presentation preference for a nicer layout
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.PresentationPreference = true;

        // Save the workbook as an HTML file using the configured options
        string htmlFile = "output.html";
        workbook.Save(htmlFile, saveOptions);
    }
}