using System;
using Aspose.Cells;

class HtmlProcessingExample
{
    static void Main()
    {
        // Load the workbook from an existing HTML file
        Workbook workbook = new Workbook("input.html");

        // Modify a cell (e.g., set value of cell B2)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["B2"].PutValue("Modified Value");

        // Configure HTML save options to disable external CSS (use inline styles only)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.DisableCss = true;

        // Save the workbook back to HTML with the specified options
        workbook.Save("output.html", saveOptions);
    }
}