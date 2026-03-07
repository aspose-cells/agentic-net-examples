using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Paths for the source XLSX file and the target HTML file
        string sourcePath = "input.xlsx";
        string targetPath = "output.html";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Create HTML save options (customize as needed)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Export the entire workbook (set to true to export only the active sheet)
        htmlOptions.ExportActiveWorksheetOnly = false;

        // Save the workbook as an HTML file using the specified options
        workbook.Save(targetPath, htmlOptions);
    }
}