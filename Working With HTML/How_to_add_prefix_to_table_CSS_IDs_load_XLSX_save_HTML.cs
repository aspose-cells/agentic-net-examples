using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and set the table CSS ID prefix
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.TableCssId = "myTablePrefix-"; // Prefix applied to tr, td, th, etc.

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}