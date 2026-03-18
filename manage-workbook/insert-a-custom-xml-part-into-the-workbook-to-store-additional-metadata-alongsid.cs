using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace CustomXmlPartDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Define the XML content and optional schema for the custom part
            string xmlData = "<MyMetadata><Author>John Doe</Author><Version>1.0</Version></MyMetadata>";
            string xmlSchema = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                               "<xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\">" +
                               "  <xs:element name=\"MyMetadata\">" +
                               "    <xs:complexType>" +
                               "      <xs:sequence>" +
                               "        <xs:element name=\"Author\" type=\"xs:string\"/>" +
                               "        <xs:element name=\"Version\" type=\"xs:string\"/>" +
                               "      </xs:sequence>" +
                               "    </xs:complexType>" +
                               "  </xs:element>" +
                               "</xs:schema>";

            // Convert strings to UTF‑8 byte arrays
            byte[] dataBytes = Encoding.UTF8.GetBytes(xmlData);
            byte[] schemaBytes = Encoding.UTF8.GetBytes(xmlSchema);

            // Add the custom XML part to the workbook (rule: CustomXmlPartCollection.Add)
            int partIndex = workbook.CustomXmlParts.Add(dataBytes, schemaBytes);

            // Optionally retrieve the added part to verify its ID or modify it further
            CustomXmlPart addedPart = workbook.CustomXmlParts[partIndex];
            Console.WriteLine($"Custom XML part added at index {partIndex}, ID: {addedPart.ID}");

            // Save the workbook (lifecycle: save)
            string outputPath = "WorkbookWithCustomXmlPart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");

            // Load the workbook back (lifecycle: load) and display the count of custom XML parts
            Workbook loadedWorkbook = new Workbook(outputPath);
            Console.WriteLine($"Number of custom XML parts after reload: {loadedWorkbook.CustomXmlParts.Count}");
        }
    }
}