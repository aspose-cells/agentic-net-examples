// Title: Assign a Unique GUID Name to an OLE Object with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, insert an OLE object using a tiny PNG, generate a GUID, set the OleObject.Name property to that GUID for uniqueness, ensure the output folder exists, and save the file as an .xlsx document.
// Keywords: Aspose.Cells OLE object name | C# set OleObject.Name | GUID OleObject property | add OLE object Aspose.Cells | unique identifier Excel OLE | .NET Excel OLE naming
// Common Searches: how to set OleObject.Name in Aspose.Cells C# | assign GUID to OLE object Aspose.Cells | unique OLE object name Excel .NET | Aspose.Cells add OLE object with custom name | C# generate unique name for Excel OLE object
// Developer Intent: Set the OleObject.Name property to a unique identifier after inserting the OLE object.
// Use Cases: Insert a single OLE object and give it a GUID‑based name so it can be referenced later without collision. | Loop through multiple OLE insertions, assigning a new Guid to each OleObject.Name to guarantee distinct identifiers. | Save a workbook with uniquely named OLE objects and retrieve them by name for further manipulation or reporting.
// AI Prompts: Generate C# code that adds an OLE object to a worksheet using Aspose.Cells and assigns a Guid string to OleObject.Name. | Show an example of inserting several OLE objects in a loop, each receiving a unique name via Guid.NewGuid(). | Explain how to locate an OLE object by its Name property after reopening the workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, insert an OLE object using a tiny PNG, generate a GUID, set the OleObject.Name property to that GUID for uniqueness, ensure the output folder exists, and save the file as an .xlsx document.
class SetOleObjectNameDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Minimal 1x1 transparent PNG (base64 encoded).
            const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2V8AAAAASUVORK5CYII=";
            byte[] imageData = Convert.FromBase64String(pngBase64);

            // Add an OLE object with the placeholder image.
            int oleIndex = worksheet.OleObjects.Add(5, 2, 200, 150, imageData);
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Assign a unique name to the OLE object.
            oleObject.Name = "OleObject_" + Guid.NewGuid().ToString("N");

            // Define output path and ensure directory exists.
            string outputPath = "OleObjectWithUniqueName.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
