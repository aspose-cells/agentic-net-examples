using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapRootElementDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define an XML schema that contains a root element
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            // Write the schema to a temporary file
            string tempSchemaPath = Path.GetTempFileName();
            File.WriteAllText(tempSchemaPath, xmlSchema);

            try
            {
                // Add the XML map to the workbook; the method returns the map index
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempSchemaPath);

                // Retrieve the XmlMap object using the index
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // Get the root element name of the XML map
                string rootElementName = xmlMap.RootElementName;

                // Output the root element name
                Console.WriteLine("Root Element Name: " + rootElementName);
            }
            finally
            {
                // Clean up the temporary file
                if (File.Exists(tempSchemaPath))
                {
                    File.Delete(tempSchemaPath);
                }
            }
        }
    }
}