using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook wb = new Workbook();

        // Access the first worksheet
        Worksheet sheet = wb.Worksheets[0];

        // Populate cells with data and apply distinct styles
        for (int i = 0; i < 10; i++)
        {
            Cell cell = sheet.Cells[i, 0];
            cell.PutValue($"Item {i + 1}");

            // Create a style for each cell
            Style style = wb.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 10 + i;
            style.Font.IsBold = (i % 2 == 0);
            cell.SetStyle(style);
        }

        // Delete rows to leave some styles unused (simulates real‑world cleanup)
        sheet.Cells.DeleteRows(5, 5);

        // Prepare custom XML data and its schema
        string xmlData = "<MyXmlData xmlns=\"http://my.namespace.com\"/>";
        string xmlSchema = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                         + "<ds:datastoreItem ds:itemID=\"{GUID}\" xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\">"
                         + "<ds:schemaRefs>"
                         + "<ds:schemaRef ds:uri=\"http://my.namespace.com\"/>"
                         + "</ds:schemaRefs>"
                         + "</ds:datastoreItem>";

        byte[] dataBytes = Encoding.UTF8.GetBytes(xmlData);
        byte[] schemaBytes = Encoding.UTF8.GetBytes(xmlSchema);

        // Add the custom XML part to the workbook (rule: CustomXmlPartCollection.Add)
        int xmlPartIndex = wb.CustomXmlParts.Add(dataBytes, schemaBytes);

        // Optionally assign a unique ID to the added part
        wb.CustomXmlParts[xmlPartIndex].ID = Guid.NewGuid().ToString();

        // Remove all unused styles in one step (rule: Workbook.RemoveUnusedStyles)
        wb.RemoveUnusedStyles();

        // Save the workbook (lifecycle save)
        wb.Save("CombinedOutput.xlsx");
    }
}