using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put the external URL into a cell (this will be the linked content source)
        string externalUrl = "https://example.com/data";
        sheet.Cells["A1"].PutValue(externalUrl);

        // Add a custom document property that links to the cell containing the URL
        // Property name: "SourceUrl", source: cell A1
        workbook.CustomDocumentProperties.AddLinkToContent("SourceUrl", "A1");

        // (Optional) Update linked properties to ensure the value is synchronized
        workbook.CustomDocumentProperties.UpdateLinkedPropertyValue();

        // Save the workbook
        workbook.Save("WorkbookWithSourceUrl.xlsx");
    }
}