using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class SetWorkbookComments
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

        // Set a multiline comment using line breaks
        properties.Comments = "This workbook was generated programmatically."
                            + Environment.NewLine
                            + "It contains sample data for demonstration purposes."
                            + Environment.NewLine
                            + "Author: John Doe"
                            + Environment.NewLine
                            + "Date: " + DateTime.Now.ToString("yyyy-MM-dd");

        // Save the workbook to a file
        workbook.Save("WorkbookWithComments.xlsx");
    }
}