using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class SetLinksUpToDate
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Mark hyperlinks as up‑to‑date
        workbook.BuiltInDocumentProperties.LinksUpToDate = true;

        // Add a sample hyperlink (optional demonstration)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

        // Save the workbook
        workbook.Save("LinksUpToDateDemo.xlsx");
    }
}