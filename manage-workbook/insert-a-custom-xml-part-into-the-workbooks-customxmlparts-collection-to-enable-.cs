using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsCustomXmlPartDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Prepare XML data and optional schema
            string xmlData = "<root><item>Sample Data</item></root>";
            string schemaData = "<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>" +
                                "<xs:element name='root'><xs:complexType>" +
                                "<xs:sequence><xs:element name='item' type='xs:string'/></xs:sequence>" +
                                "</xs:complexType></xs:element></xs:schema>";

            // Convert strings to UTF-8 byte arrays
            byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlData);
            byte[] schemaBytes = Encoding.UTF8.GetBytes(schemaData);

            // Add the custom XML part to the workbook's collection
            // Using the Add(byte[], byte[]) method as defined in the rule set
            int partIndex = workbook.CustomXmlParts.Add(xmlBytes, schemaBytes);

            // Optionally set a custom ID for the part
            CustomXmlPart part = workbook.CustomXmlParts[partIndex];
            part.ID = Guid.NewGuid().ToString();

            // Save the workbook to demonstrate that the custom XML part is stored
            workbook.Save("CustomXmlPartDemo.xlsx");
        }
    }
}