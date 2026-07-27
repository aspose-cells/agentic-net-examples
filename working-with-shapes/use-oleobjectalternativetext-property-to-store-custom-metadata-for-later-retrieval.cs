// Title: Embed and Retrieve JSON Metadata in an Excel OLE Object via AlternativeText – Aspose.Cells for .NET
// Description: Demonstrates how to add an OLE object with a minimal placeholder image to a workbook, store a JSON string in the OleObject.AlternativeText property, save the file, reload it, and read back the embedded metadata using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | OleObject AlternativeText | store JSON metadata | retrieve OLE object data | C# Excel OLE example | embed custom data in Excel | AlternativeText limit | placeholder image for OLE
// Common Searches: how to store JSON in OleObject AlternativeText Aspose.Cells | retrieve custom metadata from OLE object in Excel C# | Aspose.Cells embed data in OLE object | AlternativeText usage for OLE objects .NET | save and load OLE object metadata with Aspose.Cells
// Developer Intent: The developer needs to attach custom JSON metadata to an OLE object via the AlternativeText property and later extract it without opening the embedded file.
// Use Cases: Link a document ID to an embedded file for downstream processing. | Attach audit information (author, version) to an OLE‑embedded chart. | Persist configuration settings for a linked resource inside the workbook.
// AI Prompts: Generate C# code that serializes a dictionary to JSON, assigns it to OleObject.AlternativeText, saves the workbook, then deserializes the JSON back into a typed object using Aspose.Cells. | Explain how to handle AlternativeText length limits when storing large JSON payloads in an OLE object. | Show an example of encrypting JSON metadata before setting OleObject.AlternativeText and decrypting it after retrieval.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for OleObject

namespace OleObjectMetadataDemo
{
    // Demonstrates how to add an OLE object with a minimal placeholder image to a workbook, store a JSON string in the OleObject.AlternativeText property, save the file, reload it, and read back the embedded metadata using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path for the workbook
                string filePath = "OleObjectMetadataDemo.xlsx";

                // -------------------- Create and configure workbook --------------------
                Workbook workbook = new Workbook();                     // create a new workbook
                Worksheet sheet = workbook.Worksheets[0];              // get the first worksheet

                // Minimal 1x1 PNG image (transparent) required by Aspose.Cells for OLE objects
                // This avoids the need for System.Drawing dependencies.
                byte[] placeholderImage = Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X5eUAAAAASUVORK5CYII=");

                // Add an OLE object to the worksheet
                int oleIndex = sheet.OleObjects.Add(5, 2, 150, 150, placeholderImage);
                OleObject ole = sheet.OleObjects[oleIndex];

                // Store custom metadata in the AlternativeText property (e.g., JSON string)
                string customMetadata = "{\"DocumentId\":\"12345\",\"Author\":\"John Doe\",\"Version\":2}";
                ole.AlternativeText = customMetadata;

                // Save the workbook
                workbook.Save(filePath, SaveFormat.Xlsx);

                // -------------------- Load workbook and retrieve metadata --------------------
                if (File.Exists(filePath))
                {
                    Workbook loadedWorkbook = new Workbook(filePath); // load the saved workbook
                    OleObject loadedOle = loadedWorkbook.Worksheets[0].OleObjects[0];

                    // Retrieve the stored metadata
                    string retrievedMetadata = loadedOle.AlternativeText;

                    // Output the metadata to the console
                    Console.WriteLine("Retrieved OleObject AlternativeText:");
                    Console.WriteLine(retrievedMetadata);
                }
                else
                {
                    Console.WriteLine($"Error: The file '{filePath}' was not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
