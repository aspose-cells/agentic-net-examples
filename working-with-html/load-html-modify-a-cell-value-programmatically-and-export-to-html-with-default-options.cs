using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source HTML file into a workbook
        Workbook workbook = new Workbook("input.html");

        // Modify a cell value (e.g., cell B2 in the first worksheet)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["B2"].PutValue("Modified Value");

        // Save the workbook back to HTML using default HtmlSaveOptions
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(); // default settings
        workbook.Save("output.html", saveOptions);
    }
}