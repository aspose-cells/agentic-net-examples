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

        // Place the external URL in cell A1
        string externalUrl = "https://www.example.com/data";
        sheet.Cells["A1"].PutValue(externalUrl);

        // Add a link‑to‑content custom property named "SourceUrl" that points to cell A1
        workbook.CustomDocumentProperties.AddLinkToContent("SourceUrl", "A1");

        // Refresh the linked property value so it reflects the cell content
        workbook.CustomDocumentProperties.UpdateLinkedPropertyValue();

        // Save the workbook
        workbook.Save("LinkToContentProperty.xlsx");
    }
}