// Title: Add and Retrieve a Custom XML Part in Aspose.Cells (.NET) Using a GUID
// Description: Shows how to create a Workbook with Aspose.Cells for .NET, embed a custom XML part from a UTF‑8 byte array, assign a GUID as its identifier, save and reload the file, and fetch the same part with CustomXmlParts.SelectByID. Useful for persisting XML metadata inside Excel workbooks.
// Keywords: Aspose.Cells | custom XML part | SelectByID | GUID | C# .NET | embed XML in Excel | retrieve XML part | Workbook.CustomXmlParts | store metadata in Excel | Aspose.Cells example
// Common Searches: Aspose.Cells add custom XML part C# | SelectByID custom XML part Aspose.Cells | Assign GUID to custom XML part Aspose.Cells | Store XML metadata in Excel using Aspose.Cells | Retrieve custom XML part by ID .NET
// Developer Intent: Embed a GUID‑identified custom XML part in a workbook and later retrieve it by that ID.
// Use Cases: Persist configuration or supplemental data inside an Excel file for downstream processing. | Exchange schema‑less XML payloads alongside workbook content and access them on demand. | Link external systems to a workbook by saving their XML representation as a uniquely identifiable part.
// AI Prompts: Generate C# code that adds multiple custom XML parts with optional schemas to a workbook and retrieves each by its GUID. | Explain how to modify the XML content of an existing custom XML part identified by a GUID in Aspose.Cells. | Show how to list all custom XML parts in a workbook, displaying their IDs and data sizes.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsCustomXmlDemo
{
    // Shows how to create a Workbook with Aspose.Cells for .NET, embed a custom XML part from a UTF‑8 byte array, assign a GUID as its identifier, save and reload the file, and fetch the same part with CustomXmlParts.SelectByID. Useful for persisting XML metadata inside Excel workbooks.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Sample XML data to store in the custom XML part
            string xmlContent = "<root><item>Sample Data</item></root>";
            byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlContent);

            // Add the custom XML part (no schema data provided)
            int partIndex = workbook.CustomXmlParts.Add(xmlBytes, null);

            // Retrieve the newly added part
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

            // Output verification information
            Console.WriteLine("Original ID: " + partId);
            Console.WriteLine("Retrieved ID: " + (retrievedPart != null ? retrievedPart.ID : "Not found"));
            if (retrievedPart != null)
            {
                string retrievedXml = Encoding.UTF8.GetString(retrievedPart.Data);
                Console.WriteLine("Retrieved XML Content: " + retrievedXml);
            }
        }
    }
}
