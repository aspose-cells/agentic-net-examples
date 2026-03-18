using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and set the TableCssId prefix
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.TableCssId = "myPrefix_";

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", saveOptions);
    }
}