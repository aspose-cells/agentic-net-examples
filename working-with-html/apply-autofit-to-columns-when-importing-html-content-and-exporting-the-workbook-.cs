using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create HTML load options and enable auto‑fit for columns and rows
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
        loadOptions.AutoFitColsAndRows = true;

        // Load the HTML file using the specified options
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Export the workbook to XLSX format
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}