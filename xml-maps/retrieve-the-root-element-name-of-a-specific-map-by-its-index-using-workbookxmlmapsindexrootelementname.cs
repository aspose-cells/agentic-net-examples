// Title: How to retrieve an XML map's root element name by index using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get root element name of XML map by index | Workbook.Worksheets.XmlMaps retrieve root element name after adding XSD | Aspose.Cells read XmlMap.RootElementName from a workbook | How to obtain XML map root element in a .xlsx file with Aspose.Cells | Get XML map index and root element name in C# using Aspose.Cells
// Tags: Aspose.Cells XmlMap RootElementName | C# retrieve XML map root element | Workbook.XmlMaps add XSD schema | Aspose.Cells save workbook with XML map | XmlMap index access in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapRootElement
{
    // The example creates a temporary XSD schema, adds it to a new Workbook via Workbook.Worksheets.XmlMaps.Add, captures the returned map index, accesses the XmlMap object with Workbook.Worksheets.XmlMaps[index], reads its RootElementName property, prints the index and root element name, and saves the workbook containing the XML map.
    class Program
    {
        static void Main(string[] args)
        {
            // Prepare a simple XSD schema that defines a root element.
            string xsdContent = @"<?xml version='1.0' encoding='utf-8'?>
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
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

            // Create a temporary file for the XSD schema.
            string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeCellsXmlMapDemo");
            Directory.CreateDirectory(tempFolder);
            string xsdPath = Path.Combine(tempFolder, "catalog.xsd");
            File.WriteAllText(xsdPath, xsdContent);

            // Create a new workbook.
            using (Workbook workbook = new Workbook())
            {
                try
                {
                    // Add the XML map using the XSD file path.
                    int mapIndex = workbook.Worksheets.XmlMaps.Add(xsdPath);

                    // Retrieve the XmlMap object.
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                    // Get the root element name.
                    string rootElementName = xmlMap.RootElementName;

                    Console.WriteLine($"XML Map Index: {mapIndex}");
                    Console.WriteLine($"Root Element Name: {rootElementName}");

                    // Optional: Save the workbook to verify that the map is stored.
                    string outputPath = Path.Combine(tempFolder, "WorkbookWithXmlMap.xlsx");
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("An error occurred:");
                    Console.WriteLine(ex.Message);
                }
            }

            // Clean up temporary XSD file (optional, keep workbook file).
            // File.Delete(xsdPath);
        }
    }
}
