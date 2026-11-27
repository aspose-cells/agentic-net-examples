// Title: How to get the root element name of the first XML map in a workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get root element name from first XmlMap | retrieve root element of XML map added to workbook with Aspose.Cells | C# read XmlMap.RootElementName after adding XSD schema in Aspose.Cells | example code for accessing first XmlMap root element in an Excel workbook | how to use Aspose.Cells XmlMaps collection to obtain root element name
// Tags: Aspose.Cells XmlMap RootElementName | C# retrieve first XmlMap root element | Aspose.Cells add XML map from XSD | Excel workbook XmlMaps collection access | read root element name from XSD schema using Aspose.Cells

using System;
using Aspose.Cells;

// The example creates an empty Workbook, adds an XML map from an XSD schema, accesses the first XmlMap in the workbook's XmlMaps collection, reads its RootElementName property, and prints the root element name to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // Define a simple XML schema (XSD) as a string.
            // This schema has a root element named "root".
            string xsd = @"<?xml version='1.0' encoding='utf-8'?>
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='root'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='item' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='Name' type='xs:string'/>
              <xs:element name='Value' type='xs:int'/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

            // Add the XML map to the workbook.
            // The Add method returns the index of the newly added map.
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xsd);

            // Verify that at least one XmlMap exists.
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps were added to the workbook.");
                return;
            }

            // Retrieve the first XmlMap (index 0) – it is the same map we just added.
            XmlMap firstMap = workbook.Worksheets.XmlMaps[0];

            // Use the RootElementName property to get the name of the root element.
            string rootElementName = firstMap.RootElementName;

            // Output the result.
            Console.WriteLine("Root element name of the first XML map: " + rootElementName);
        }
        catch (Exception ex)
        {
            // Any unexpected errors are reported here.
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
