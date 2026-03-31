using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsCustomXmlDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create a new workbook --------------------
            Workbook workbook = new Workbook(); // creates an empty workbook

            // -------------------- Add a custom XML part --------------------
            // Sample XML data and optional schema
            string xmlData = "<root><item>Original Value</item></root>";
            string xmlSchema = "<xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\"><xs:element name=\"root\"/></xs:schema>";

            // Convert strings to UTF-8 byte arrays as required by Add method
            byte[] dataBytes = Encoding.UTF8.GetBytes(xmlData);
            byte[] schemaBytes = Encoding.UTF8.GetBytes(xmlSchema);

            // Add the custom XML part to the workbook; Add returns the index of the new part
            int partIndex = workbook.CustomXmlParts.Add(dataBytes, schemaBytes);

            // Retrieve the added part using the indexer
            CustomXmlPart customPart = workbook.CustomXmlParts[partIndex];

            // Display the automatically generated ID
            Console.WriteLine("Initial Part ID: " + customPart.ID);

            // -------------------- Manipulate the XML content --------------------
            // Read current XML content
            string currentXml = Encoding.UTF8.GetString(customPart.Data);
            Console.WriteLine("Current XML Data: " + currentXml);

            // Modify the XML data (e.g., change the item value)
            string modifiedXml = "<root><item>Modified Value</item></root>";
            customPart.Data = Encoding.UTF8.GetBytes(modifiedXml);

            // Optionally, update the schema (here we replace it with a new simple schema)
            string newSchema = "<xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\"><xs:element name=\"root\"><xs:complexType><xs:sequence><xs:element name=\"item\" type=\"xs:string\"/></xs:sequence></xs:complexType></xs:element></xs:schema>";
            customPart.SchemaData = Encoding.UTF8.GetBytes(newSchema);

            // Change the ID to a known GUID for later retrieval
            string newId = Guid.NewGuid().ToString();
            customPart.ID = newId;
            Console.WriteLine("Updated Part ID: " + customPart.ID);

            // -------------------- Save the workbook --------------------
            string filePath = "CustomXmlDemo.xlsx";
            workbook.Save(filePath); // uses the provided Save method

            // -------------------- Load the workbook and verify changes --------------------
            Workbook loadedWorkbook = new Workbook(filePath); // uses the provided load constructor

            // Retrieve the part by the known ID
            CustomXmlPart loadedPart = loadedWorkbook.CustomXmlParts.SelectByID(newId);
            if (loadedPart != null)
            {
                string loadedXml = Encoding.UTF8.GetString(loadedPart.Data);
                Console.WriteLine("Loaded XML Data: " + loadedXml);
                Console.WriteLine("Loaded Schema (as string): " + Encoding.UTF8.GetString(loadedPart.SchemaData));
            }
            else
            {
                Console.WriteLine("Custom XML part with ID not found after reload.");
            }
        }
    }
}