using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add sample data to the first worksheet
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("Item");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["A2"].PutValue("Apple");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["A3"].PutValue("Banana");
        ws.Cells["B3"].PutValue(20);

        // Apply distinct styles to create styles that may become unused
        Style style1 = wb.CreateStyle();
        style1.Font.Name = "Arial";
        style1.Font.Size = 12;
        ws.Cells["A2"].SetStyle(style1);

        Style style2 = wb.CreateStyle();
        style2.Font.Name = "Times New Roman";
        style2.Font.Size = 14;
        style2.Font.IsBold = true;
        ws.Cells["A3"].SetStyle(style2);

        // Delete a row to make style1 unused
        ws.Cells.DeleteRow(2); // Removes the row containing style1

        // Prepare custom XML data and schema
        string xmlData = "<MyData xmlns=\"http://example.com\"><Item>Sample</Item></MyData>";
        string xmlSchema = "<xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\"><xs:element name=\"MyData\"/></xs:schema>";
        byte[] dataBytes = Encoding.UTF8.GetBytes(xmlData);
        byte[] schemaBytes = Encoding.UTF8.GetBytes(xmlSchema);

        // Add the custom XML part to the workbook
        wb.CustomXmlParts.Add(dataBytes, schemaBytes);

        // Remove all unused styles in a single step
        wb.RemoveUnusedStyles();

        // Save the workbook with all changes applied
        wb.Save("ProcessedWorkbook.xlsx");
    }
}