using System;
using Aspose.Cells;

class HtmlToExcel
{
    static void Main()
    {
        // Input HTML file path
        string inputHtml = "input.html";

        // Output Excel file path
        string outputXlsx = "output.xlsx";

        // Create load options and enable auto‑fit for columns and rows
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        loadOptions.AutoFitColsAndRows = true;

        // Load the HTML file into a workbook with the specified options
        Workbook workbook = new Workbook(inputHtml, loadOptions);

        // Save the workbook as an XLSX file
        workbook.Save(outputXlsx, SaveFormat.Xlsx);
    }
}