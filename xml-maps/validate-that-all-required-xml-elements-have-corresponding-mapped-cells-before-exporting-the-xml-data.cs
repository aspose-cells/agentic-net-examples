// Title: Validate required XML elements are mapped to cells before exporting XML with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells verify required XML elements have mapped cells before ExportXml | How to use GetAreasOfXmlMapQuery to detect unmapped XML nodes in a workbook | Validate XML schema required fields are linked to worksheet cells with Aspose.Cells | ExportXml only after confirming all required XML paths are mapped in C# | Aspose.Cells example for checking missing XML map queries before export
// Tags: Validate XML map required elements Aspose.Cells | GetAreasOfXmlMapQuery C# | ExportXml after validation Aspose.Cells | XML schema mapping verification .NET | Worksheet XML map cell area check

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace XmlMapValidationDemo
{
    // The sample creates a temporary XSD schema and matching XML file, adds an XML map to a new Aspose.Cells workbook, imports the XML data, defines required XPath element paths, and uses GetAreasOfXmlMapQuery to ensure each required element has at least one mapped cell area. If all required elements are mapped, ExportXml writes the XML file; otherwise, missing elements are reported and the export is aborted. Temporary files are cleaned up after execution.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for temporary files
            string schemaPath = "schema.xsd";
            string xmlPath = "data.xml";
            string outputXmlPath = "exported.xml";

            // Create a simple XSD schema defining required elements
            string xsdContent = @"<?xml version='1.0' encoding='utf-8'?>
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Data'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Item' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='Name' type='xs:string'/>   <!-- required -->
              <xs:element name='Value' type='xs:integer'/> <!-- required -->
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

            // Create a simple XML file that matches the schema
            string xmlContent = @"<?xml version='1.0' encoding='utf-8'?>
<Data>
  <Item>
    <Name>Item1</Name>
    <Value>100</Value>
  </Item>
  <Item>
    <Name>Item2</Name>
    <Value>200</Value>
  </Item>
</Data>";

            // Write schema and XML to disk (Aspose.Cells expects file paths)
            File.WriteAllText(schemaPath, xsdContent);
            File.WriteAllText(xmlPath, xmlContent);

            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Add the XML map based on the XSD schema
                int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // 3. Import XML data into the first worksheet (creates the mapping)
                workbook.ImportXml(xmlPath, "Sheet1", 0, 0);

                // 4. Retrieve the worksheet where data was imported
                Worksheet worksheet = workbook.Worksheets["Sheet1"];

                // 5. Define the required XML element paths (XPath style without namespace)
                // These paths must match the ones used in the XSD.
                List<string> requiredPaths = new List<string>
                {
                    "/Data/Item/Name",
                    "/Data/Item/Value"
                };

                // 6. Validate that each required path has at least one mapped cell area
                List<string> missingElements = new List<string>();
                foreach (string path in requiredPaths)
                {
                    // GetAreasOfXmlMapQuery returns an array of CellArea objects that are linked to the path
                    CellArea[] areas = worksheet.GetAreasOfXmlMapQuery(path, xmlMap);
                    if (areas == null || areas.Length == 0)
                    {
                        missingElements.Add(path);
                    }
                }

                // 7. Report validation result
                if (missingElements.Count == 0)
                {
                    // All required elements are mapped – proceed with export
                    workbook.ExportXml(xmlMap.Name, outputXmlPath);
                    Console.WriteLine($"All required XML elements are mapped. Export succeeded to '{outputXmlPath}'.");
                }
                else
                {
                    Console.WriteLine("Validation failed. The following required XML elements have no mapped cells:");
                    foreach (string missing in missingElements)
                    {
                        Console.WriteLine($"  - {missing}");
                    }
                    Console.WriteLine("Export aborted.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Clean up temporary files (optional)
                if (File.Exists(schemaPath)) File.Delete(schemaPath);
                if (File.Exists(xmlPath)) File.Delete(xmlPath);
            }
        }
    }
}
