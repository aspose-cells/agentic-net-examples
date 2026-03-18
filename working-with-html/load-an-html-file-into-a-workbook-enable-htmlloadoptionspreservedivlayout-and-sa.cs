using System;
using Aspose.Cells;

class HtmlToExcel
{
    static void Main()
    {
        // Create HTML load options and enable support for <div> layout
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        loadOptions.SupportDivTag = true;

        // Load the HTML file into a workbook using the specified options
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Save the loaded workbook as an XLSX file
        workbook.Save("output.xlsx");
    }
}