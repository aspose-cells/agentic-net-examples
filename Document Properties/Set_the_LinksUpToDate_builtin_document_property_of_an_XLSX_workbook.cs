using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class SetLinksUpToDateDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the LinksUpToDate built‑in document property
        workbook.BuiltInDocumentProperties.LinksUpToDate = true;

        // Add a sample hyperlink (optional, just to have a link in the workbook)
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

        // Save the workbook to a file
        workbook.Save("LinksUpToDateDemo.xlsx");
    }
}