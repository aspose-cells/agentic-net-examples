// Title: Import XML into linked cells with Workbook.ImportXml after adding an XML map (Aspose.Cells for .NET C#)
// Description: Demonstrates how to create a workbook, add an XSD‑based XML map, link cell A1 to the "/Root/Item" path, import XML data with Workbook.ImportXml, save the result, and clean up temporary files.
// Keywords: Aspose.Cells | C# | .NET | ImportXml | XML map | LinkToXmlMap | linked cells | XSD schema | XML import to Excel | Excel automation | data migration
// Common Searches: Aspose.Cells import XML linked cells C# | Workbook.ImportXml after adding XML map | LinkToXmlMap example .NET | How to map XML to Excel cells using Aspose | Import XML into Excel template with Aspose.Cells
// Developer Intent: Add an XML map, link worksheet cells to schema paths, import XML data, and generate a populated XLSX file.
// Use Cases: Populate a template where each <Item> element fills successive rows in column A. | Migrate XML‑based configuration data into predefined Excel reports. | Automate financial models by linking specific XML elements to calculation cells.
// AI Prompts: Show how to map multiple XML elements to different columns using linked cells. | Provide a version that reads the XSD and XML from streams instead of files. | Explain handling of XML namespaces when linking cells with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook, add an XSD‑based XML map, link cell A1 to the "/Root/Item" path, import XML data with Workbook.ImportXml, save the result, and clean up temporary files.
class ImportXmlWithLinkedCells
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Prepare a simple XML schema (XSD) and XML data for the demo
        string xsdContent = @"<?xml version='1.0' encoding='utf-8'?>
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Root'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Item' type='xs:string' maxOccurs='unbounded'/>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

        string xmlContent = @"<?xml version='1.0' encoding='utf-8'?>
<Root>
  <Item>First</Item>
  <Item>Second</Item>
  <Item>Third</Item>
</Root>";

        // Write the schema and XML to temporary files
        string schemaPath = Path.Combine(Path.GetTempPath(), "sample.xsd");
        string xmlPath = Path.Combine(Path.GetTempPath(), "sample.xml");
        File.WriteAllText(schemaPath, xsdContent);
        File.WriteAllText(xmlPath, xmlContent);

        // Add the XML map to the workbook using the schema file
        int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
        xmlMap.Name = "SampleMap";

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Link cell A1 to the XML map path "/Root/Item"
        // This creates a linked cell that will be populated when XML is imported
        sheet.Cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Item");

        // Import the XML data into the worksheet starting at cell A1
        // The ImportXml method will fill the linked cells according to the map
        workbook.ImportXml(xmlPath, sheet.Name, 0, 0);

        // Save the resulting workbook
        string outputPath = "LinkedXmlOutput.xlsx";
        workbook.Save(outputPath);

        // Clean up temporary files
        File.Delete(schemaPath);
        File.Delete(xmlPath);

        Console.WriteLine($"Workbook saved to '{outputPath}'. Linked cells have been populated from XML.");
    }
}
