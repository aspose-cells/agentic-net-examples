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

            // Define custom XML metadata
            string xmlContent = "<metadata><author>John Doe</author><created>2024-01-01</created></metadata>";
            byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlContent);

            // Add the custom XML part (no schema data in this example)
            int partIndex = workbook.CustomXmlParts.Add(xmlBytes, null);

            // Retrieve the added part and assign a unique identifier
            CustomXmlPart customPart = workbook.CustomXmlParts[partIndex];
            string uniqueId = Guid.NewGuid().ToString();
            customPart.ID = uniqueId;

            // Save the workbook containing the custom XML part
            string filePath = "CustomXmlDemo.xlsx";
            workbook.Save(filePath);

            // Load the workbook from disk
            Workbook loadedWorkbook = new Workbook(filePath);

            // Retrieve the custom XML part using its unique identifier
            CustomXmlPart retrievedPart = loadedWorkbook.CustomXmlParts.SelectByID(uniqueId);

            // Output the retrieved information
            Console.WriteLine("Retrieved Part ID: " + (retrievedPart != null ? retrievedPart.ID : "Not found"));
            if (retrievedPart != null)
            {
                string retrievedXml = Encoding.UTF8.GetString(retrievedPart.Data);
                Console.WriteLine("Retrieved XML Content: " + retrievedXml);
            }
        }
    }
}