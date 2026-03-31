using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Mark all hyperlinks as up‑to‑date
        workbook.BuiltInDocumentProperties.LinksUpToDate = true;

        // Add a sample hyperlink (optional, just to illustrate)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

        // Save the workbook
        workbook.Save("LinksUpToDateDemo.xlsx");
    }
}