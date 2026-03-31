using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class SetLinksUpToDateFalse
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

        // Flag hyperlinks as outdated by setting LinksUpToDate to false
        builtInProps.LinksUpToDate = false;

        // Add a sample hyperlink (optional, just for demonstration)
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");

        // Save the workbook to an XLSX file
        workbook.Save("LinksOutdated.xlsx");
    }
}