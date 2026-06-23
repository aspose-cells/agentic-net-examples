using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Define a simple XML schema (XSD) for the map
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Root'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Item' type='xs:string' />
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                    </xs:schema>";

                // Add the XML map to the workbook; returns the map index
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);

                // Retrieve the map and get its root element name
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                string rootElementName = xmlMap?.RootElementName ?? "Unknown";

                Console.WriteLine("Root Element Name of the first XML map: " + rootElementName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}