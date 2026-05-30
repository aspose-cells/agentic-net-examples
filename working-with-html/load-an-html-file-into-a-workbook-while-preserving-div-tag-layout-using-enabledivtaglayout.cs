using System;
using Aspose.Cells;

class LoadHtmlWithDivTag
{
    static void Main()
    {
        // Create HTML load options and enable DIV tag layout support
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
        loadOptions.SupportDivTag = true;

        // Load the HTML file into a workbook using the specified options
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Example: read a cell value to verify loading
        Worksheet worksheet = workbook.Worksheets[0];
        Console.WriteLine("Cell A1 value: " + worksheet.Cells["A1"].StringValue);

        // Save the workbook to an Excel file
        workbook.Save("output.xlsx");
    }
}