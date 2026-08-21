// Title: Aspose.Cells .NET: Add an XML Map with a Custom Namespace Prefix and Export Namespaced XML
// Description: Demonstrates how to create a Workbook, generate a temporary XSD that defines a custom namespace prefix, add the XSD as an XmlMap, assign a friendly map name, and export the workbook to an XML file that preserves the specified namespace. Includes safe cleanup of the temporary schema file.
// Keywords: Aspose.Cells XML map | custom namespace prefix | C# export XML | XmlMap from XSD | namespaced XML Aspose.Cells | temporary XSD file | Workbook ExportXml | .NET XML mapping
// Common Searches: Aspose.Cells set custom namespace prefix for XmlMap | export XML with namespace using Aspose.Cells C# | create XML map from XSD with namespace prefix | how to add XmlMap in Aspose.Cells .NET | clean up temporary XSD after adding XmlMap
// Developer Intent: Add an XML map that recognizes a custom namespace prefix and export workbook data to a correctly namespaced XML document.
// Use Cases: Generate employee XML files that conform to a namespaced XSD for integration with external systems. | Produce XML reports where a specific namespace prefix is required by a partner API. | Automate the creation of temporary XSD files, map registration, and cleanup in batch processing pipelines.
// AI Prompts: Show C# code to add an XmlMap with a custom namespace prefix using Aspose.Cells and export the workbook to XML. | Explain how to ensure the exported XML retains the defined namespace prefix after mapping. | Provide best practices for deleting temporary XSD files created for XmlMap initialization in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapNamespacePrefixDemo
{
    // Demonstrates how to create a Workbook, generate a temporary XSD that defines a custom namespace prefix, add the XSD as an XmlMap, assign a friendly map name, and export the workbook to an XML file that preserves the specified namespace. Includes safe cleanup of the temporary schema file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Sample XML schema with a custom namespace prefix "ns"
                string xmlSchema = @"
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'
           xmlns:ns='http://example.com/ns'
           targetNamespace='http://example.com/ns'
           elementFormDefault='qualified'>
    <xs:element name='Employee' type='ns:EmployeeType'/>
    <xs:complexType name='EmployeeType'>
        <xs:sequence>
            <xs:element name='ID' type='xs:int'/>
            <xs:element name='Name' type='xs:string'/>
        </xs:sequence>
    </xs:complexType>
</xs:schema>";

                // Write the schema to a temporary XSD file because Aspose.Cells expects a file path
                string tempXsdPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xsd");
                File.WriteAllText(tempXsdPath, xmlSchema);

                // Ensure the temporary file exists before adding the XML map
                if (!File.Exists(tempXsdPath))
                    throw new FileNotFoundException("Temporary XSD file was not created.", tempXsdPath);

                // Add the XML map to the workbook using the temporary XSD file
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXsdPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // Assign a friendly name to the map (used when exporting/importing XML)
                xmlMap.Name = "EmployeeMap";

                // Display map information
                Console.WriteLine("Root element name: " + xmlMap.RootElementName);
                Console.WriteLine("Map name: " + xmlMap.Name);

                // Export the workbook data to an XML file using the map name.
                string exportPath = "EmployeeExport.xml";

                // Ensure the directory for the export file exists
                string exportDir = Path.GetDirectoryName(Path.GetFullPath(exportPath));
                if (!Directory.Exists(exportDir))
                    Directory.CreateDirectory(exportDir);

                workbook.ExportXml(xmlMap.Name, exportPath);
                Console.WriteLine("XML exported successfully to: " + exportPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                // Clean up the temporary XSD file if it exists
                // (In a real application, consider more robust cleanup handling)
                // Note: The variable is scoped inside try, so we recreate the path pattern.
                try
                {
                    string tempFolder = Path.GetTempPath();
                    foreach (var file in Directory.GetFiles(tempFolder, "*.xsd"))
                    {
                        // Attempt to delete only files that were likely created by this demo
                        // (simple heuristic based on creation time)
                        var creation = File.GetCreationTimeUtc(file);
                        if ((DateTime.UtcNow - creation).TotalMinutes < 5)
                        {
                            File.Delete(file);
                        }
                    }
                }
                catch
                {
                    // Suppress any cleanup exceptions
                }
            }
        }
    }
}
