using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

        // Set a multiline comment describing the workbook
        properties.Comments =
            "This workbook contains financial data.\n" +
            "Generated on: " + DateTime.Now.ToString("yyyy-MM-dd") + "\n" +
            "Prepared by: John Doe.\n" +
            "Notes: Review the summary sheet for key metrics.";

        // Save the workbook to a file
        workbook.Save("WorkbookWithComments.xlsx");
    }
}