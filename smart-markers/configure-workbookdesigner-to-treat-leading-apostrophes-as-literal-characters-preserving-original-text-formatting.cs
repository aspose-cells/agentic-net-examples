using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a new workbook (template) and a worksheet
        // -------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Place a smart marker where the value will be inserted.
        // The marker references the "Text" property of the data source named "Data".
        sheet.Cells["A1"].PutValue("&Data.Text");

        // -------------------------------------------------
        // 2. Configure the workbook to keep leading apostrophes
        //    as literal characters (not as formatting).
        // -------------------------------------------------
        // When QuotePrefixToStyle is false, a leading single quote is stored
        // as part of the cell value instead of being treated as a style flag.
        workbook.Settings.QuotePrefixToStyle = false;

        // -------------------------------------------------
        // 3. Prepare a data source that contains a leading apostrophe
        // -------------------------------------------------
        var dataSource = new
        {
            Text = "'Value that starts with an apostrophe"
        };

        // -------------------------------------------------
        // 4. Initialise WorkbookDesigner, bind the data source and process
        // -------------------------------------------------
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource("Data", dataSource);
        designer.Process(); // populate the smart marker

        // -------------------------------------------------
        // 5. Save the resulting workbook
        // -------------------------------------------------
        workbook.Save("Result.xlsx");
    }
}