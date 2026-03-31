using System;
using System.IO;
using Aspose.Cells;

namespace XmlMapAssociationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define a simple XML schema
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
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

            // Write the schema to a temporary file
            string tempXsdPath = Path.GetTempFileName();
            File.WriteAllText(tempXsdPath, xmlSchema);

            // Add the XML map to the workbook
            int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXsdPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "ItemMap";

            // Clean up the temporary file
            File.Delete(tempXsdPath);

            // Get the first worksheet and its cells collection
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (A2:B4)
            cells["A2"].PutValue("Item1");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("Item2");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("Item3");
            cells["B4"].PutValue(30);

            // Associate cells with the XML map using indexed paths
            for (int row = 1; row <= 3; row++) // rows 2 to 4 (zero‑based)
            {
                int itemIndex = row; // 1‑based index for XML items
                cells.LinkToXmlMap(xmlMap.Name, row, 0, $"/Root/Item[{itemIndex}]/Name");
                cells.LinkToXmlMap(xmlMap.Name, row, 1, $"/Root/Item[{itemIndex}]/Value");
            }

            // Save the workbook
            workbook.Save("MappedWorkbook.xlsx");
        }
    }
}