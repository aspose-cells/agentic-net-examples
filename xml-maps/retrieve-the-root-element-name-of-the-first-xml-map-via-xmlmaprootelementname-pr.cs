using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapRootDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define a simple XML schema (XSD) with a root element
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' type='xs:string' />
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            // Write the schema to a temporary file
            string tempFilePath = Path.GetTempFileName();
            File.WriteAllText(tempFilePath, xmlSchema);

            // Add the XML map to the workbook; Add returns the index of the new map
            int mapIndex = workbook.Worksheets.XmlMaps.Add(tempFilePath);

            // Retrieve the first XmlMap (index 0) and get its root element name
            XmlMap firstMap = workbook.Worksheets.XmlMaps[0];
            string rootName = firstMap.RootElementName;

            // Output the root element name
            Console.WriteLine("Root element name of the first XML map: " + rootName);

            // Clean up the temporary file
            File.Delete(tempFilePath);
        }
    }
}