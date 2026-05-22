using System;
using Aspose.Cells;

class ImportHtmlTable
{
    static void Main()
    {
        // Create an empty workbook
        Workbook workbook = new Workbook();

        // Configure HTML load options
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();

        // Define a table load option:
        // - Import the first table (index 0) from the HTML
        // - Convert it to a ListObject (structured table) so column headers are recognized
        HtmlTableLoadOption tableOption = new HtmlTableLoadOption
        {
            TableIndex = 0,
            TableToListObject = true
        };

        // Add the option to the collection
        loadOptions.TableLoadOptions.Add(tableOption);

        // Load the HTML file with the specified options
        workbook = new Workbook("input.html", loadOptions);

        // Save the workbook as an Excel file
        workbook.Save("output.xlsx");
    }
}