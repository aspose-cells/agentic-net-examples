using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsCustomXmlDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided Workbook constructor)
            Workbook workbook = new Workbook();

            // Sample XML data representing a simple book catalog
            string xmlData = @"
                <catalog xmlns=""http://example.com/bookcatalog"">
                    <book id=""b1"">
                        <title>Learning Aspose.Cells</title>
                        <author>John Doe</author>
                        <price>29.99</price>
                    </book>
                    <book id=""b2"">
                        <title>Advanced .NET Programming</title>
                        <author>Jane Smith</author>
                        <price>39.99</price>
                    </book>
                </catalog>";

            // Corresponding XSD schema for the book catalog
            string schemaData = @"
                <?xml version=""1.0"" encoding=""UTF-8""?>
                <xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema""
                           targetNamespace=""http://example.com/bookcatalog""
                           xmlns=""http://example.com/bookcatalog""
                           elementFormDefault=""qualified"">
                  <xs:element name=""catalog"">
                    <xs:complexType>
                      <xs:sequence>
                        <xs:element name=""book"" maxOccurs=""unbounded"">
                          <xs:complexType>
                            <xs:sequence>
                              <xs:element name=""title"" type=""xs:string""/>
                              <xs:element name=""author"" type=""xs:string""/>
                              <xs:element name=""price"" type=""xs:decimal""/>
                            </xs:sequence>
                            <xs:attribute name=""id"" type=""xs:string"" use=""required""/>
                          </xs:complexType>
                        </xs:element>
                      </xs:sequence>
                    </xs:complexType>
                  </xs:element>
                </xs:schema>";

            // Convert XML and schema strings to UTF‑8 byte arrays
            byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlData);
            byte[] schemaBytes = Encoding.UTF8.GetBytes(schemaData);

            // Add the custom XML part (uses the provided CustomXmlParts.Add method)
            int partIndex = workbook.CustomXmlParts.Add(xmlBytes, schemaBytes);

            // Optionally, add a content type property to describe the custom XML part
            // (uses the provided ContentTypeProperties.Add method)
            workbook.ContentTypeProperties.Add("BookCatalogPart", "Added", "string");

            // Save the workbook (uses the provided Save method)
            workbook.Save("BookCatalogWorkbook.xlsx");
        }
    }
}