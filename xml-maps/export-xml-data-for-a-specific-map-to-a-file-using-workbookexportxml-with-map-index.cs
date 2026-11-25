// Title: Export data from an Aspose.Cells workbook to an XML file using a specific XML map in C#
// Common Searches: Aspose.Cells C# export workbook data to XML using a defined XML map | Workbook.ExportXml method with map index example in .NET | How to add an XML map from an XSD and export XML with Aspose.Cells | Export Excel worksheet to XML file using Aspose.Cells XML map name | C# code to write temporary XSD, create XML map and export XML with Aspose
// Tags: Workbook.ExportXml map name parameter | create XML map from schema file Aspose.Cells | temporary schema file for XML map Aspose | Aspose.Cells XML map data export C# | export workbook to XML via specific map index

using System;
using System.IO;
using Aspose.Cells;

namespace ExportXmlMapDemo
{
    // The example creates a temporary XSD schema, adds it as an XML map to a new workbook, fills sample data, and then uses Workbook.ExportXml with the map's name to write the mapped data to an XML file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define output XML file path
            string outputXmlPath = Path.Combine(Directory.GetCurrentDirectory(), "ExportedData.xml");

            // Temporary XSD file for the XML map
            string tempXsdPath = Path.Combine(Path.GetTempPath(), "PeopleMap.xsd");

            try
            {
                // 1. Prepare a simple XML schema (XSD) and write it to a temporary file
                string xmlSchema = @"
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='People'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Person' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='Id' type='xs:int'/>
              <xs:element name='Name' type='xs:string'/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

                File.WriteAllText(tempXsdPath, xmlSchema);

                // 2. Create a new workbook and populate sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                cells["A1"].PutValue("Id");
                cells["B1"].PutValue("Name");
                cells["A2"].PutValue(1);
                cells["B2"].PutValue("Alice");
                cells["A3"].PutValue(2);
                cells["B3"].PutValue("Bob");

                // 3. Add the XML map using the temporary XSD file
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXsdPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // 4. Export the XML data using the map's name
                workbook.ExportXml(xmlMap.Name, outputXmlPath);

                Console.WriteLine($"XML data exported successfully to: {outputXmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the export process:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                // Clean up the temporary XSD file if it exists
                if (File.Exists(tempXsdPath))
                {
                    try
                    {
                        File.Delete(tempXsdPath);
                    }
                    catch
                    {
                        // Ignored – best effort cleanup
                    }
                }
            }
        }
    }
}
