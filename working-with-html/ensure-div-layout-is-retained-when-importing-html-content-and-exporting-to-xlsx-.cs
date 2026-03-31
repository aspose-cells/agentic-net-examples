using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Enable support for <div> tag layout when loading HTML.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
        loadOptions.SupportDivTag = true;

        // Load the HTML file into a workbook while preserving the DIV layout.
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Export the workbook to an XLSX file.
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}