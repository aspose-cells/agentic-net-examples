using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set an initial value in cell A1 (the source for the linked property)
        sheet.Cells["A1"].PutValue("Initial Content");

        // Add a custom document property that links to the content of cell A1
        workbook.CustomDocumentProperties.AddLinkToContent("LinkedProp", "A1");

        // Update the linked property value from the cell content
        workbook.CustomDocumentProperties.UpdateLinkedPropertyValue();

        // Change the cell value to demonstrate updating the linked property later
        sheet.Cells["A1"].PutValue("Updated Content");

        // Refresh the linked property value again after the cell change
        workbook.CustomDocumentProperties.UpdateLinkedPropertyValue();

        // Save the workbook as an XLSX file
        workbook.Save("LinkedContentProperty.xlsx");
    }
}