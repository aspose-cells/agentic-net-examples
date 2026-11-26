// Title: Assign a new GUID to an OleObject's ClassIdentifier in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set custom ClassId for embedded OLE object | how to change the ClassIdentifier of an OleObject in a .xlsx file using Aspose.Cells | generate and assign new GUID to OleObject ClassIdentifier with Aspose.Cells .NET | update OLE object associated application by modifying ClassIdentifier in Aspose.Cells workbook
// Tags: set OleObject ClassIdentifier GUID Aspose.Cells | embed OLE object with custom ClassId .NET | modify OLE object ClassIdentifier in Excel file | Aspose.Cells OLE object GUID assignment | change OLE object associated application programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, embeds a text file as an OLE object with a placeholder PNG image, generates a new GUID, assigns its byte array to the OleObject.ClassIdentifier to alter the linked application, and saves the result as an .xlsx file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Prepare a temporary folder for demo files
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeOleDemo");
        Directory.CreateDirectory(tempFolder);

        // Create a simple text file that will be embedded as the OLE object data
        string embeddedFilePath = Path.Combine(tempFolder, "sample.txt");
        File.WriteAllText(embeddedFilePath, "This is sample embedded content.");
        byte[] objectData = File.ReadAllBytes(embeddedFilePath);

        // Minimal PNG image (1x1 pixel) to be used as the picture displayed for the OLE object
        byte[] imageData = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
            0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
            0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
            0x42,0x60,0x82
        };

        // Add an OleObject to the worksheet
        int oleIndex = worksheet.OleObjects.Add(5, 5, 100, 100, imageData);
        OleObject oleObject = worksheet.OleObjects[oleIndex];

        // Set the embedded object's data and display options
        oleObject.ObjectData = objectData;
        oleObject.DisplayAsIcon = true;
        oleObject.ImageSourceFullName = "icon.png"; // optional visual name

        // Generate a new GUID and assign it to the ClassIdentifier (ClassId) of the OleObject
        Guid newGuid = Guid.NewGuid();
        byte[] guidBytes = newGuid.ToByteArray(); // 16‑byte representation
        oleObject.ClassIdentifier = guidBytes;

        // Save the workbook
        string outputPath = Path.Combine(tempFolder, "OleObjectWithNewGuid.xlsx");
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook saved to: {outputPath}");
        Console.WriteLine($"New ClassIdentifier GUID assigned to OleObject: {newGuid}");
    }
}
