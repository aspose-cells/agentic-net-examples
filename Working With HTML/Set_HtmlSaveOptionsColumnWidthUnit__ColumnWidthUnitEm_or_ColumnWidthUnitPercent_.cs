using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Create HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}