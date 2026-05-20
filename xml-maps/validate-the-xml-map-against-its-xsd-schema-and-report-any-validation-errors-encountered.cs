using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using Aspose.Cells;

class ValidateXmlMap
{
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

            // Sample XML that will be validated against the XSD
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

            // Create a new workbook (no external template file required)
            Workbook workbook = new Workbook();

            // Add the XML map using the XSD schema
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xsd);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "DataMap";

            // Prepare a schema set for validation
            XmlSchemaSet schemaSet = new XmlSchemaSet();
            using (XmlReader schemaReader = XmlReader.Create(new StringReader(xsd)))
            {
                schemaSet.Add(null, schemaReader);
            }

            // Collect validation errors
            List<string> errors = new List<string>();
            XmlReaderSettings settings = new XmlReaderSettings
            {
                Schemas = schemaSet,
                ValidationType = System.Xml.ValidationType.Schema // disambiguated
            };
            settings.ValidationEventHandler += (sender, e) => errors.Add(e.Message);

            // Perform validation
            using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml), settings))
            {
                while (xmlReader.Read()) { }
            }

            // Output validation results
            if (errors.Count == 0)
            {
                Console.WriteLine("XML is valid against the XSD schema.");
            }
            else
            {
                Console.WriteLine("Validation errors encountered:");
                foreach (string err in errors)
                {
                    Console.WriteLine("- " + err);
                }
            }
        }
        catch (Exception ex)
        {
            // General exception handling to avoid runtime crashes
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}