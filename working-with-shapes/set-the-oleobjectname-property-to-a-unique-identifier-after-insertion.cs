// Title: How to Assign a Unique GUID Name to an OLE Object with Aspose.Cells for .NET
// Description: Demonstrates creating a workbook, inserting an OLE object with a 1×1 PNG placeholder, generating a GUID, setting the OleObject.Name property to that GUID, optionally assigning a label, and saving the file as XLSX using Aspose.Cells for C#.
// Keywords: Aspose.Cells OLE object name | set OleObject.Name C# | GUID for OLE object | placeholder image OLE Aspose | Aspose.Cells add OLE object | unique OLE object identifier | OleObject.Label property
// Common Searches: Aspose.Cells set unique name for OLE object | C# assign GUID to OleObject.Name | add OLE object with placeholder image Aspose.Cells | how to label OLE objects in Excel using Aspose | retrieve OLE object by name Aspose.Cells
// Developer Intent: Assign a globally unique name to an OleObject immediately after it is added to a worksheet.
// Use Cases: Insert several OLE objects and give each a distinct GUID‑based name for later lookup or automation. | Generate Excel reports where embedded objects must be uniquely identifiable for downstream processing. | Add OLE objects with a temporary image and a user‑friendly label while preserving a unique internal name.
// AI Prompts: Write C# code that adds multiple OLE objects to a worksheet with Aspose.Cells and assigns each a unique name using Guid.NewGuid(). | Show how to find an OleObject in a saved workbook by its Name property with Aspose.Cells for .NET. | Explain how to replace the placeholder image of an existing OleObject without changing its assigned GUID name.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectNameDemo
{
    // Demonstrates creating a workbook, inserting an OLE object with a 1×1 PNG placeholder, generating a GUID, setting the OleObject.Name property to that GUID, optionally assigning a label, and saving the file as XLSX using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a simple placeholder image (1x1 pixel PNG) in memory
                byte[] placeholderImage = CreatePlaceholderImage();

                // Add an OLE object with the placeholder image
                int oleIndex = worksheet.OleObjects.Add(5, 2, 200, 150, placeholderImage);
                OleObject oleObject = worksheet.OleObjects[oleIndex];

                // Assign a unique name to the OLE object
                oleObject.Name = "OleObject_" + Guid.NewGuid().ToString("N");

                // Optionally set a label for display purposes
                oleObject.Label = "Sample OLE Object";

                // Ensure the output directory exists
                string outputPath = "OleObjectWithUniqueName.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Returns a 1x1 transparent PNG image as a byte array
        private static byte[] CreatePlaceholderImage()
        {
            // Base64 representation of a 1x1 transparent PNG
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X0ZcAAAAASUVORK5CYII=";
            return Convert.FromBase64String(base64Png);
        }
    }
}
