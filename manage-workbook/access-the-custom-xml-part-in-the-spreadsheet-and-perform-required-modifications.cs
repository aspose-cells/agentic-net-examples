using System;
using System.Text;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsCustomXmlDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (lifecycle rule: use Workbook(string) constructor)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the collection of custom XML parts
            CustomXmlPartCollection customXmlParts = workbook.CustomXmlParts;

            // Ensure there is at least one custom XML part; add one if none exist
            if (customXmlParts.Count == 0)
            {
                // Sample XML data to add
                string initialXml = "<root><item>Original</item></root>";
                byte[] xmlBytes = Encoding.UTF8.GetBytes(initialXml);

                // Add the XML part (rule: CustomXmlPartCollection.Add)
                int newIndex = customXmlParts.Add(xmlBytes, null);
                Console.WriteLine($"Added new custom XML part at index {newIndex}");
            }

            // Retrieve the first custom XML part
            CustomXmlPart xmlPart = customXmlParts[0];

            // Convert the existing XML data to a string for manipulation
            string xmlContent = Encoding.UTF8.GetString(xmlPart.Data);
            Console.WriteLine("Original XML Data:");
            Console.WriteLine(xmlContent);

            // Load the XML into XDocument for safe modification
            XDocument doc = XDocument.Parse(xmlContent);

            // Example modification: change the value of <item> element to "Modified"
            XElement itemElement = doc.Root.Element("item");
            if (itemElement != null)
            {
                itemElement.Value = "Modified";
            }

            // Convert the modified XML back to byte[] and assign it back to the part (rule: CustomXmlPart.Data)
            string modifiedXml = doc.Declaration != null ? doc.Declaration + doc.ToString() : doc.ToString();
            xmlPart.Data = Encoding.UTF8.GetBytes(modifiedXml);

            Console.WriteLine("Modified XML Data:");
            Console.WriteLine(modifiedXml);

            // Optionally, update the schema data (example: set a simple schema)
            string schemaXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
                <xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
                    <xs:element name=""root"">
                        <xs:complexType>
                            <xs:sequence>
                                <xs:element name=""item"" type=""xs:string""/>
                            </xs:sequence>
                        </xs:complexType>
                    </xs:element>
                </xs:schema>";
            xmlPart.SchemaData = Encoding.UTF8.GetBytes(schemaXml);

            // Save the workbook with the updated custom XML part (lifecycle rule: Workbook.Save)
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved as output.xlsx");
        }
    }
}