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
            // Create a new workbook (uses the Workbook() constructor)
            Workbook wb = new Workbook();

            // Sample XML data and an optional XML schema
            string xmlData = "<Employee><Name>John Doe</Name><Id>123</Id></Employee>";
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Employee'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Name' type='xs:string'/>
                                                <xs:element name='Id' type='xs:int'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                 </xs:schema>";

            // Convert strings to UTF‑8 byte arrays
            byte[] dataBytes = Encoding.UTF8.GetBytes(xmlData);
            byte[] schemaBytes = Encoding.UTF8.GetBytes(xmlSchema);

            // Add the custom XML part to the workbook (returns the index of the new part)
            int partIndex = wb.CustomXmlParts.Add(dataBytes, schemaBytes);

            // Retrieve the added part by index
            CustomXmlPart part = wb.CustomXmlParts[partIndex];

            // Display the automatically generated ID
            Console.WriteLine($"Added Custom XML Part ID: {part.ID}");

            // Assign a custom ID for easier retrieval later
            string customId = Guid.NewGuid().ToString();
            part.ID = customId;
            Console.WriteLine($"Custom ID set to: {customId}");

            // Modify the XML content of the part
            string updatedXml = "<Employee><Name>Jane Smith</Name><Id>456</Id></Employee>";
            part.Data = Encoding.UTF8.GetBytes(updatedXml);
            Console.WriteLine("Custom XML part data updated.");

            // Save the workbook to disk (uses Workbook.Save(string))
            string filePath = "CustomXmlDemo.xlsx";
            wb.Save(filePath);
            Console.WriteLine($"Workbook saved to '{filePath}'.");

            // Load the workbook from the saved file (uses Workbook(string) constructor)
            Workbook loadedWb = new Workbook(filePath);

            // Retrieve the custom XML part by the custom ID
            CustomXmlPart retrievedPart = loadedWb.CustomXmlParts.SelectByID(customId);
            if (retrievedPart != null)
            {
                string retrievedXml = Encoding.UTF8.GetString(retrievedPart.Data);
                Console.WriteLine("Retrieved Custom XML Part Data:");
                Console.WriteLine(retrievedXml);
            }
            else
            {
                Console.WriteLine("Custom XML part with the specified ID was not found.");
            }

            // Display total count of custom XML parts in the loaded workbook
            Console.WriteLine($"Total Custom XML Parts in loaded workbook: {loadedWb.CustomXmlParts.Count}");
        }
    }
}