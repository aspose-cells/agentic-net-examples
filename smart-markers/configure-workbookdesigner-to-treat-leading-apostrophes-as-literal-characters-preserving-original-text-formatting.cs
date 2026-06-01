using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (template)
        Workbook workbook = new Workbook();

        // Example cell containing a leading apostrophe
        workbook.Worksheets[0].Cells["A1"].PutValue("'SampleText");

        // Configure the workbook so that leading apostrophes are treated as literal characters
        // (i.e., they are kept in the cell value and not applied as QuotePrefix style)
        workbook.Settings.QuotePrefixToStyle = false;

        // Initialize the WorkbookDesigner with the prepared workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Process any smart markers (none in this simple example)
        designer.Process();

        // Save the workbook; the cell will retain the leading apostrophe as part of its value
        workbook.Save("Output.xlsx");
    }
}