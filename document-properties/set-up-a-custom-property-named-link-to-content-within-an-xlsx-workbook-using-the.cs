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

        // Put a sample value into cell A1
        sheet.Cells["A1"].PutValue("Sample Content");

        // Add a custom document property that links directly to cell A1
        workbook.CustomDocumentProperties.AddLinkToContent("link to content", "Sheet1!A1");

        // Synchronize the linked property value with the cell content
        workbook.CustomDocumentProperties.UpdateLinkedPropertyValue();

        // Save the workbook as an XLSX file
        workbook.Save("LinkedContentProperty.xlsx");
    }
}