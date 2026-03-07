using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create HTML load options (self‑closing tags are handled automatically)
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();

        // Load the HTML file into a workbook using the specified options
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Save the loaded workbook as an XLSX file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}