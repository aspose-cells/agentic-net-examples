using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsCustomXmlFreezeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare custom XML data and optional schema
            string xmlData = "<MyData xmlns=\"http://example.com\"><Item>Value</Item></MyData>";
            string xmlSchema = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                               "<xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" targetNamespace=\"http://example.com\" xmlns=\"http://example.com\" elementFormDefault=\"qualified\">" +
                               "<xs:element name=\"MyData\"><xs:complexType><xs:sequence><xs:element name=\"Item\" type=\"xs:string\"/></xs:sequence></xs:complexType></xs:element>" +
                               "</xs:schema>";

            // Convert strings to UTF‑8 byte arrays
            byte[] dataBytes = Encoding.UTF8.GetBytes(xmlData);
            byte[] schemaBytes = Encoding.UTF8.GetBytes(xmlSchema);

            // Add the custom XML part to the workbook
            CustomXmlPartCollection customParts = workbook.CustomXmlParts;
            int partIndex = customParts.Add(dataBytes, schemaBytes);

            // Retrieve the added part to obtain its ID (optional, for reference)
            CustomXmlPart addedPart = customParts[partIndex];
            string partId = addedPart.ID;

            // Write the part ID into the worksheet so the user can see the reference
            sheet.Cells["A1"].PutValue("Custom XML Part ID:");
            sheet.Cells["B1"].PutValue(partId);

            // Freeze the first row (which contains the reference) for easy navigation
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // Row and column are zero‑based indexes of the cell where the split occurs.
            // To freeze the top row, split at row index 1 (second row) and freeze 1 row.
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook
            workbook.Save("CustomXmlWithFreeze.xlsx");
        }
    }
}