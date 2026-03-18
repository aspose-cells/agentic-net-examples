using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

class CustomXmlPartDemo
{
    static void Main()
    {
        // Create a new workbook (uses the provided Workbook constructor)
        Workbook wb = new Workbook();

        // Sample XML data and its schema
        string xmlData = "<MyData xmlns=\"http://example.com\"><Item>Value</Item></MyData>";
        string xmlSchema = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                         + "<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema' targetNamespace='http://example.com' xmlns='http://example.com' elementFormDefault='qualified'>"
                         + "<xs:element name='MyData'><xs:complexType><xs:sequence><xs:element name='Item' type='xs:string'/></xs:sequence></xs:complexType></xs:element>"
                         + "</xs:schema>";

        // Convert strings to UTF‑8 byte arrays (required by Add method)
        byte[] dataBytes = Encoding.UTF8.GetBytes(xmlData);
        byte[] schemaBytes = Encoding.UTF8.GetBytes(xmlSchema);

        // Add a custom XML part to the workbook; Add returns the index of the new part
        int partIndex = wb.CustomXmlParts.Add(dataBytes, schemaBytes);

        // Retrieve the newly added CustomXmlPart instance using the indexer
        CustomXmlPart part = wb.CustomXmlParts[partIndex];

        // Assign a unique identifier to the part (useful for later lookup)
        part.ID = Guid.NewGuid().ToString();

        // Update the XML content of the part
        string updatedXml = "<MyData xmlns=\"http://example.com\"><Item>UpdatedValue</Item></MyData>";
        part.Data = Encoding.UTF8.GetBytes(updatedXml);

        // Save the workbook (uses the provided Save method)
        wb.Save("CustomXmlPartDemo.xlsx");
    }
}