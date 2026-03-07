using System;
using Aspose.Cells;

namespace AsposeCellsHtmlAutoFitDemo
{
    class Program
    {
        static void Main()
        {
            // Create HTML load options and enable auto‑fit for columns and rows
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.AutoFitColsAndRows = true;

            // Load the HTML file with the specified options
            Workbook workbook = new Workbook("input.html", loadOptions);

            // Save the loaded workbook as an XLSX file
            workbook.Save("output.xlsx");
        }
    }
}