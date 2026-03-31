using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create HTML load options and specify the format
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
        // Enable automatic fitting of both columns and rows during HTML import
        loadOptions.AutoFitColsAndRows = true;

        // Load the HTML file with the specified options
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Save the loaded workbook as an XLSX file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}