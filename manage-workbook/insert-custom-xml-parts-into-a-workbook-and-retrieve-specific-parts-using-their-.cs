using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsCustomXmlDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Prepare first XML data (no schema)
            string xmlData1 = "<root><item>First</item></root>";
            byte[] dataBytes1 = Encoding.UTF8.GetBytes(xmlData1);

            // Add first custom XML part and obtain its index
            int index1 = workbook.CustomXmlParts.Add(dataBytes1, null);
            // Retrieve the part and assign a unique ID
            CustomXmlPart part1 = workbook.CustomXmlParts[index1];
            string id1 = Guid.NewGuid().ToString();
            part1.ID = id1;

            // Prepare second XML data with a simple schema
            string xmlData2 = "<root><value>Second</value></root>";
            string schemaData2 = "<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'><xs:element name='root'><xs:complexType><xs:sequence><xs:element name='value' type='xs:string'/></xs:sequence></xs:complexType></xs:element></xs:schema>";
            byte[] dataBytes2 = Encoding.UTF8.GetBytes(xmlData2);
            byte[] schemaBytes2 = Encoding.UTF8.GetBytes(schemaData2);

            // Add second custom XML part
            int index2 = workbook.CustomXmlParts.Add(dataBytes2, schemaBytes2);
            CustomXmlPart part2 = workbook.CustomXmlParts[index2];
            string id2 = Guid.NewGuid().ToString();
            part2.ID = id2;

            // Save the workbook to disk
            string filePath = "CustomXmlDemo.xlsx";
            workbook.Save(filePath);

            // Load the workbook back
            Workbook loadedWorkbook = new Workbook(filePath);

            // Retrieve parts by their IDs
            CustomXmlPart retrievedPart1 = loadedWorkbook.CustomXmlParts.SelectByID(id1);
            CustomXmlPart retrievedPart2 = loadedWorkbook.CustomXmlParts.SelectByID(id2);

            // Output the IDs and XML content to verify
            Console.WriteLine("Retrieved Part 1 ID: " + (retrievedPart1?.ID ?? "Not found"));
            Console.WriteLine("Retrieved Part 1 XML: " + Encoding.UTF8.GetString(retrievedPart1?.Data ?? new byte[0]));

            Console.WriteLine("Retrieved Part 2 ID: " + (retrievedPart2?.ID ?? "Not found"));
            Console.WriteLine("Retrieved Part 2 XML: " + Encoding.UTF8.GetString(retrievedPart2?.Data ?? new byte[0]));
        }
    }
}