// Title: Setting a custom namespace prefix for an XML map and importing namespaced XML into an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add XML map with custom namespace prefix from XSD | Import XML document that uses a namespace prefix into Excel with Aspose.Cells | Set XML map root element namespace handling in Aspose.Cells .NET | Create workbook and map namespaced XML schema using Aspose.Cells API | Auto‑fit columns after importing namespaced XML with Aspose.Cells
// Tags: xmlmap custom namespace prefix aspose.cells | add xml map from xsd schema c# | import xml data into worksheet aspose.cells | auto fit columns after xml import aspose.cells | save workbook as xlsx aspose.cells

using System;
using System.IO;
using Aspose.Cells;

namespace XmlMapNamespacePrefixDemo
{
    // The example creates a new workbook, defines an XSD schema with a custom namespace prefix, adds it as an XML map, assigns a map name, imports a matching namespaced XML document into the worksheet starting at cell A1, auto‑fits the columns, and saves the file as an XLSX workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output Excel file
            const string outputFile = "NamespacedXmlMapDemo.xlsx";

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // 2. Define an XSD schema that uses a custom namespace with a prefix (e.g., "ns")
                //    The schema defines a root element <ns:Catalog> containing multiple <ns:Product> items.
                string xsdSchema = @"
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'
           targetNamespace='http://example.com/catalog'
           xmlns:ns='http://example.com/catalog'
           elementFormDefault='qualified'>

  <xs:element name='Catalog'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Product' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='ID' type='xs:int'/>
              <xs:element name='Name' type='xs:string'/>
              <xs:element name='Price' type='xs:decimal'/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>

</xs:schema>";

                // 3. Add the XML map to the workbook using the XSD string.
                //    The Add method returns the index of the newly created map.
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xsdSchema);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // 4. Set a custom name for the map (optional, but demonstrates the Name property).
                //    The map name is independent of the namespace prefix.
                xmlMap.Name = "CatalogMap";

                // 5. Display the root element name and confirm the namespace handling.
                Console.WriteLine($"XmlMap added. Name: {xmlMap.Name}");
                Console.WriteLine($"Root element name (without prefix): {xmlMap.RootElementName}");

                // 6. Prepare an XML document that uses the same namespace and prefix defined in the XSD.
                string xmlData = @"
<ns:Catalog xmlns:ns='http://example.com/catalog'>
    <ns:Product>
        <ns:ID>1</ns:ID>
        <ns:Name>Apple</ns:Name>
        <ns:Price>0.99</ns:Price>
    </ns:Product>
    <ns:Product>
        <ns:ID>2</ns:ID>
        <ns:Name>Banana</ns:Name>
        <ns:Price>0.59</ns:Price>
    </ns:Product>
</ns:Catalog>";

                // 7. Import the XML data into the worksheet starting at cell A1.
                //    ImportXml will automatically map the XML elements according to the XmlMap,
                //    respecting the namespace prefix defined in the schema.
                workbook.ImportXml(xmlData, sheet.Name, 0, 0);

                // 8. Optionally, auto‑fit columns for better visibility.
                sheet.AutoFitColumns();

                // 9. Save the workbook.
                workbook.Save(outputFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the XML map:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
