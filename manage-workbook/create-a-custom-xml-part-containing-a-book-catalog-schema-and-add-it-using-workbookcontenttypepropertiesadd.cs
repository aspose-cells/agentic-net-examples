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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Sample XML data representing a simple book catalog
            string xmlData = @"
<catalog xmlns=""http://example.com/bookcatalog"">
    <book id=""bk101"">
        <author>Gambardella, Matthew</author>
        <title>XML Developer's Guide</title>
        <genre>Computer</genre>
        <price>44.95</price>
        <publish_date>2000-10-01</publish_date>
        <description>An in-depth look at creating applications with XML.</description>
    </book>
</catalog>";

            // Corresponding XML schema (XSD) for the catalog
            string xmlSchema = @"
<?xml version=""1.0"" encoding=""UTF-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"" targetNamespace=""http://example.com/bookcatalog"" xmlns=""http://example.com/bookcatalog"" elementFormDefault=""qualified"">
  <xs:element name=""catalog"">
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""book"" maxOccurs=""unbounded"">
          <xs:complexType>
            <xs:sequence>
              <xs:element name=""author"" type=""xs:string""/>
              <xs:element name=""title"" type=""xs:string""/>
              <xs:element name=""genre"" type=""xs:string""/>
              <xs:element name=""price"" type=""xs:decimal""/>
              <xs:element name=""publish_date"" type=""xs:date""/>
              <xs:element name=""description"" type=""xs:string""/>
            </xs:sequence>
            <xs:attribute name=""id"" type=""xs:string"" use=""required""/>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

            // Convert XML and schema strings to UTF-8 byte arrays
            byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlData);
            byte[] schemaBytes = Encoding.UTF8.GetBytes(xmlSchema);

            // Add the custom XML part to the workbook using the provided API
            int partIndex = workbook.CustomXmlParts.Add(xmlBytes, schemaBytes);

            // Optional: retrieve the part to verify its ID or content
            CustomXmlPart addedPart = workbook.CustomXmlParts[partIndex];
            Console.WriteLine($"Custom XML part added at index {partIndex}, ID: {addedPart.ID}");

            // Save the workbook to a file
            string outputPath = "BookCatalogWorkbook.xlsx";
            workbook.Save(outputPath);

            // Load the workbook back to confirm the custom XML part persists
            Workbook loadedWorkbook = new Workbook(outputPath);
            Console.WriteLine($"Loaded workbook contains {loadedWorkbook.CustomXmlParts.Count} custom XML part(s).");
        }
    }
}