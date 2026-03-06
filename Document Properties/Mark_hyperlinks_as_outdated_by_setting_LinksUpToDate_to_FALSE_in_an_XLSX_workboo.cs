using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a sample hyperlink (optional, just to have a link in the file)
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");

        // Mark hyperlinks as outdated
        workbook.BuiltInDocumentProperties.LinksUpToDate = false;

        // Save the workbook to an XLSX file
        workbook.Save("OutdatedLinks.xlsx");
    }
}