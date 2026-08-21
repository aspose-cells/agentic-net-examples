// Title: C# Example: Store and Retrieve JSON Metadata in an Excel OLE Object via OleObject.AlternativeText (Aspose.Cells)
// Description: Demonstrates how to add an OLE object to a worksheet, embed a JSON string as custom metadata using the OleObject.AlternativeText property, save the workbook, and later reload it to read the stored metadata. Includes a helper that creates a 1×1 PNG icon for the OLE object.
// Keywords: Aspose.Cells | C# | .NET | OleObject | AlternativeText | metadata | JSON | embed data in Excel | OLE object custom properties | sample code | GitHub example | Excel automation
// Common Searches: Aspose.Cells store JSON in OleObject.AlternativeText | retrieve custom metadata from Excel OLE object C# | how to embed data in OLE object using Aspose.Cells | AlternativeText property size limit | C# example for saving metadata in Excel shape
// Developer Intent: Persist custom JSON data inside an OLE object’s AlternativeText field and read it back after the workbook is saved.
// Use Cases: Link a unique document ID to each embedded OLE object for quick lookup. | Save author, creation date, or version info with the object to avoid external databases. | Create a lightweight, self‑contained metadata store for embedded charts, diagrams, or files.
// AI Prompts: Generate C# code that serializes any object to JSON and assigns it to OleObject.AlternativeText using Aspose.Cells. | Write a method that scans all OleObjects in a worksheet and returns a dictionary of their AlternativeText values. | Explain the character limit of the AlternativeText property and suggest ways to handle metadata that exceeds this limit.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add an OLE object to a worksheet, embed a JSON string as custom metadata using the OleObject.AlternativeText property, save the workbook, and later reload it to read the stored metadata. Includes a helper that creates a 1×1 PNG icon for the OLE object.
class OleObjectMetadataExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Generate a simple 1x1 pixel PNG image to use as the OLE object icon
            byte[] iconImage = GeneratePngIcon();

            // Add an OLE object to the worksheet at row 5, column 2 with size 150x150 pixels
            int oleIndex = sheet.OleObjects.Add(5, 2, 150, 150, iconImage);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Store custom metadata in the AlternativeText property (e.g., JSON string)
            string customMetadata = "{\"DocumentId\":12345,\"Author\":\"John Doe\",\"Created\":\"2024-01-01\"}";
            ole.AlternativeText = customMetadata;

            // Save the workbook to a file
            string filePath = "OleObjectWithMetadata.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // -------------------------------------------------
            // Later: Load the workbook and retrieve the metadata
            // -------------------------------------------------
            if (File.Exists(filePath))
            {
                Workbook loadedWorkbook = new Workbook(filePath);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

                // Assuming the OLE object is still at index 0 (first added)
                if (loadedSheet.OleObjects.Count > 0)
                {
                    OleObject loadedOle = loadedSheet.OleObjects[0];

                    // Retrieve the stored metadata
                    string retrievedMetadata = loadedOle.AlternativeText;

                    // Output the metadata to the console
                    Console.WriteLine("Retrieved metadata from OleObject.AlternativeText:");
                    Console.WriteLine(retrievedMetadata);
                }
                else
                {
                    Console.WriteLine("No OLE objects found in the loaded worksheet.");
                }
            }
            else
            {
                Console.WriteLine($"File not found: {filePath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }

    // Helper method to generate a minimal 1x1 transparent PNG image as a byte array
    private static byte[] GeneratePngIcon()
    {
        // Base64-encoded 1x1 transparent PNG
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2V8AAAAASUVORK5CYII=";
        return Convert.FromBase64String(base64Png);
    }
}
