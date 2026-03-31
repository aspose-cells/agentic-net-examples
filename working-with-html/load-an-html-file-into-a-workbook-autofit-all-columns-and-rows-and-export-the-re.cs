using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Input HTML file path
        string htmlFile = "input.html";

        // Output XLSX file path
        string xlsxFile = "output.xlsx";

        // Create HTML load options and enable auto‑fit for columns and rows
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
        loadOptions.AutoFitColsAndRows = true;

        // Load the HTML file into a workbook using the specified options
        Workbook workbook = new Workbook(htmlFile, loadOptions);

        // Save the workbook as an XLSX file
        workbook.Save(xlsxFile, SaveFormat.Xlsx);
    }
}