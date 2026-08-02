using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class AddLinkToContentCustomProperty
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the external URL to be stored
        string externalUrl = "https://www.example.com/data";

        // Place the URL in a cell (e.g., A1) – this cell will be the source for the linked property
        sheet.Cells["A1"].PutValue(externalUrl);

        // Add a custom document property named "SourceUrl" that links to the content of cell A1
        // The source parameter must be a cell reference or named range that contains the value
        workbook.CustomDocumentProperties.AddLinkToContent("SourceUrl", "A1");

        // (Optional) Update linked properties to ensure the property value reflects the cell content
        workbook.CustomDocumentProperties.UpdateLinkedPropertyValue();

        // Save the workbook to a file
        workbook.Save("WorkbookWithLinkedUrl.xlsx");
    }
}