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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Prepare XML data for the custom part
            string xmlContent = "<metadata><author>John Doe</author><created>2024-01-01</created></metadata>";
            byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlContent);

            // (Optional) Prepare an XML schema – can be null if not needed
            byte[] schemaBytes = null;

            // Add the custom XML part to the workbook and obtain its index
            int partIndex = workbook.CustomXmlParts.Add(xmlBytes, schemaBytes);

            // Retrieve the newly added part via the index
            CustomXmlPart customPart = workbook.CustomXmlParts[partIndex];

            // Assign a unique identifier (GUID) to the part
            string partId = Guid.NewGuid().ToString();
            customPart.ID = partId;

            // Save the workbook containing the custom XML part
            string filePath = "CustomXmlDemo.xlsx";
            workbook.Save(filePath);

            // Load the workbook from disk
            Workbook loadedWorkbook = new Workbook(filePath);

            // Retrieve the custom XML part using its unique ID
            CustomXmlPart retrievedPart = loadedWorkbook.CustomXmlParts.SelectByID(partId);

            // Output the ID and XML content of the retrieved part
            if (retrievedPart != null)
            {
                Console.WriteLine("Retrieved Part ID: " + retrievedPart.ID);
                string retrievedXml = Encoding.UTF8.GetString(retrievedPart.Data);
                Console.WriteLine("Retrieved XML Content: " + retrievedXml);
            }
            else
            {
                Console.WriteLine("Custom XML part with ID not found.");
            }
        }
    }
}