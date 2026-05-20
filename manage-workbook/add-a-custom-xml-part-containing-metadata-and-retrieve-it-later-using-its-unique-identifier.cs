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
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();

            // Sample XML data to store in the custom XML part
            string xmlContent = "<root><item>Sample Metadata</item></root>";
            byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlContent);

            // Add the custom XML part (no schema data in this example)
            int partIndex = workbook.CustomXmlParts.Add(xmlBytes, null);

            // Retrieve the added part and assign a unique identifier (GUID)
            CustomXmlPart customPart = workbook.CustomXmlParts[partIndex];
            string uniqueId = Guid.NewGuid().ToString();
            customPart.ID = uniqueId;

            // Save the workbook containing the custom XML part
            string filePath = "CustomXmlDemo.xlsx";
            workbook.Save(filePath);

            // ---------- Load the workbook and retrieve the custom XML part ----------
            Workbook loadedWorkbook = new Workbook(filePath);

            // Use the unique identifier to locate the custom XML part
            CustomXmlPart retrievedPart = loadedWorkbook.CustomXmlParts.SelectByID(uniqueId);

            // Output verification information
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