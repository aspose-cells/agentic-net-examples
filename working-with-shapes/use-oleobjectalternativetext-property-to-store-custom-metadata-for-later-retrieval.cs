// Title: How to embed a file as an OLE object and store custom JSON metadata using OleObject.AlternativeText in Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a text file as an OLE object and writes a JSON string to its AlternativeText property with Aspose.Cells. | Show how to open a saved Excel workbook, locate the first OLE object, and extract the AlternativeText value using Aspose.Cells for .NET. | Adapt the sample to insert a PDF file as an OLE object and store XML metadata in the AlternativeText field.
// Common Searches: Aspose.Cells embed OLE object and save custom metadata in AlternativeText | C# read AlternativeText from OLE object in Excel workbook using Aspose.Cells | store JSON string in OLE object's AlternativeText property with Aspose.Cells .NET | how to retrieve embedded file metadata from Excel using Aspose.Cells
// Tags: OLE object insertion Aspose.Cells C# | AlternativeText property custom metadata | JSON metadata storage in OLE object | AlternativeText retrieval Aspose.Cells | workbook save load with embedded OLE

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, inserts a text file as an OLE object, saves a JSON string in the object's AlternativeText property, persists the file, then reloads the workbook to read back the stored metadata.
class OleObjectMetadataExample
{
    static void Main()
    {
        try
        {
            // Prepare a sample text file to embed.
            string sourceFile = @"C:\Temp\Sample.txt";
            string sourceDir = Path.GetDirectoryName(sourceFile);
            if (!string.IsNullOrEmpty(sourceDir) && !Directory.Exists(sourceDir))
            {
                Directory.CreateDirectory(sourceDir);
            }

            if (!File.Exists(sourceFile))
            {
                File.WriteAllText(sourceFile, "Sample content");
            }

            // Read file bytes for embedding.
            byte[] oleData = File.ReadAllBytes(sourceFile);

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add an OLE object (embed the text file) at cell A1 with size 200x100 pixels.
            // Use the overload that accepts byte[] data and progId.
            int oleObjectIndex = sheet.OleObjects.Add(0, 0, 200, 100, oleData, "Package");

            // Retrieve the OLE object that was just added.
            OleObject oleObject = sheet.OleObjects[oleObjectIndex];

            // Store custom metadata in the AlternativeText property (e.g., JSON string).
            string customMetadata = "{\"DocumentId\":12345,\"Author\":\"John Doe\"}";
            oleObject.AlternativeText = customMetadata;

            // Save the workbook.
            string filePath = "OleObjectMetadata.xlsx";
            workbook.Save(filePath);

            // Load the workbook and retrieve the stored metadata.
            if (File.Exists(filePath))
            {
                Workbook loadedWorkbook = new Workbook(filePath);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                OleObject loadedOleObject = loadedSheet.OleObjects[0];
                string retrievedMetadata = loadedOleObject.AlternativeText;
                Console.WriteLine("Retrieved metadata: " + retrievedMetadata);
            }
            else
            {
                Console.WriteLine($"Workbook file not found: {filePath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
