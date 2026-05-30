using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the HTML file and enable support for <div> layout
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
        loadOptions.SupportDivTag = true;
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Modify a cell in the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Modified Value");

        // Save the workbook as HTML with inline styles only (disable external CSS)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.DisableCss = true;
        workbook.Save("output.html", saveOptions);
    }
}