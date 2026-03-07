using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create HTML load options and specify the format
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
        // Enable auto‑fit of columns and rows during import
        loadOptions.AutoFitColsAndRows = true;

        // Load the HTML file using the specified options
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Save the workbook as an XLSX file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}