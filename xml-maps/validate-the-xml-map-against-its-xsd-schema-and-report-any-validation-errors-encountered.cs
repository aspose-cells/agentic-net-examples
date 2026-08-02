using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using Aspose.Cells;

class Program
{
    // Alias to avoid ambiguity between Aspose.Cells.ValidationType and System.Xml.ValidationType
    private static readonly Type XmlValidationType = typeof(System.Xml.ValidationType);

    static void Main()
    {
        try
        {
            // XSD schema defining the XML structure
            string xsd = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                            <xs:element name='Data'>
                                <xs:complexType>
                                    <xs:sequence>
                                        <xs:element name='Item' maxOccurs='unbounded'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Name' type='xs:string'/>
                                                    <xs:element name='Value' type='xs:integer'/>
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                    </xs:sequence>
                                </xs:complexType>
                            </xs:element>
                          </xs:schema>";

            // Sample XML to be validated (contains an intentional error)
            string xml = @"<Data>
                            <Item>
                                <Name>Item1</Name>
                                <Value>100</Value>
                            </Item>
                            <Item>
                                <Name>Item2</Name>
                                <Value>InvalidInteger</Value>
                            </Item>
                          </Data>";

            // Create a workbook (no external template file used, but check if needed)
            Workbook workbook = new Workbook();

            // Add the XML map based on the XSD schema
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xsd);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "DataMap";

            // Collect validation error messages
            List<string> validationErrors = new List<string>();

            // Configure XML reader settings for schema validation
            XmlReaderSettings settings = new XmlReaderSettings
            {
                // Use fully qualified enum to avoid ambiguity
                ValidationType = System.Xml.ValidationType.Schema
            };

            // Load the XSD schema from the string
            using (StringReader xsdReader = new StringReader(xsd))
            {
                settings.Schemas.Add(null, XmlReader.Create(xsdReader));
            }

            // Capture validation events
            settings.ValidationEventHandler += (sender, args) =>
            {
                validationErrors.Add(args.Message);
            };

            // Validate the XML string against the XSD schema
            using (StringReader xmlReader = new StringReader(xml))
            using (XmlReader reader = XmlReader.Create(xmlReader, settings))
            {
                try
                {
                    while (reader.Read()) { }
                }
                catch (XmlException ex)
                {
                    validationErrors.Add("XML parsing error: " + ex.Message);
                }
            }

            // Output validation results
            if (validationErrors.Count == 0)
            {
                Console.WriteLine("XML is valid against the XSD schema.");
            }
            else
            {
                Console.WriteLine("Validation errors encountered:");
                foreach (string error in validationErrors)
                {
                    Console.WriteLine("- " + error);
                }
            }

            // Optional: export data to XML using the defined map (demonstration purpose)
            // workbook.ExportXml(xmlMap.Name, "ExportedData.xml");
        }
        catch (Exception ex)
        {
            // General exception handling to prevent runtime crashes
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }
}