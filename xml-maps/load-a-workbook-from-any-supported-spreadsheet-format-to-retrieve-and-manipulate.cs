using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XmlMapManipulationDemo
    {
        public static void Run()
        {
            // Path to the source workbook (any supported spreadsheet format)
            string sourcePath = "input.xlsx";

            // Load the workbook using the constructor that accepts a file path
            Workbook workbook = new Workbook(sourcePath);

            // Access the collection of XML maps in the workbook
            XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

            // If the workbook does not contain any XML maps, add a sample map from an XSD schema
            if (xmlMaps.Count == 0)
            {
                // Sample XSD schema defining a simple XML structure
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Root'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Item' type='xs:string'/>
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                    </xs:schema>";

                // Write the schema to a temporary file because XmlMapCollection.Add expects a file path
                string tempSchemaPath = "tempSchema.xsd";
                File.WriteAllText(tempSchemaPath, xmlSchema);

                // Add the XML map to the workbook; the method returns the index of the new map
                xmlMaps.Add(tempSchemaPath);

                // Clean up the temporary schema file
                File.Delete(tempSchemaPath);
            }

            // Retrieve the first XML map (now guaranteed to exist)
            XmlMap map = xmlMaps[0];

            // Example manipulation: rename the XML map
            map.Name = "MyCustomMap";

            // Export the XML data linked by this map to a file
            string exportPath = "exported.xml";
            workbook.ExportXml(map.Name, exportPath);

            // Optionally, import new XML data into the first worksheet using the map
            // This demonstrates how to update the workbook with external XML content
            string importXmlPath = "newData.xml";
            if (File.Exists(importXmlPath))
            {
                // Import XML into the first worksheet starting at cell A1 (row 0, column 0)
                workbook.ImportXml(importXmlPath, workbook.Worksheets[0].Name, 0, 0);
            }

            // Save the workbook after modifications, preserving the original format
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }

    public class Program
    {
        public static void Main()
        {
            XmlMapManipulationDemo.Run();
        }
    }
}