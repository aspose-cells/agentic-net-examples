// Title: Clone an OLE object to a different worksheet using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add an OLE object with a preview image and embedded Excel data, then duplicate that OLE object on another sheet while preserving its image, embedded file, and key properties, and finally save the result.
// Keywords: Aspose.Cells | C# | .NET | clone OLE object | copy OLE object worksheet | OleObjects.Add | ObjectData | FileFormatType | DisplayAsIcon | ProgID | Excel OLE embedding | sample code | GitHub example
// Common Searches: Aspose.Cells clone OLE object C# | copy OLE object to another sheet .NET | duplicate embedded Excel OLE in workbook | how to transfer OLE object properties with Aspose.Cells | sample code for cloning OLE objects in C#
// Developer Intent: Programmatically duplicate an existing OLE object and place the copy on a different worksheet while retaining its visual preview and embedded data.
// Use Cases: Create a master template with a reusable OLE chart and replicate it across multiple report sheets. | Embed a supplemental Excel file as an OLE object in a dashboard workbook and copy it to each department tab. | Maintain the snapshot image of an OLE object when reorganizing worksheets for printing or page‑break adjustments.
// AI Prompts: Generate C# code with Aspose.Cells that clones an OLE object from one worksheet to another, preserving image preview and embedded file. | Explain which OLE object properties must be copied (FileFormatType, DisplayAsIcon, ProgID) to achieve an exact duplicate in Aspose.Cells. | Provide error‑handling patterns for missing preview image or embedded file when cloning OLE objects with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add an OLE object with a preview image and embedded Excel data, then duplicate that OLE object on another sheet while preserving its image, embedded file, and key properties, and finally save the result.
class CloneOleObjectDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Source worksheet where the original OLE object resides
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Prepare image data for the OLE object's preview (snapshot)
            byte[] snapshotImage;
            const string snapshotPath = "snapshot.png";
            if (File.Exists(snapshotPath))
            {
                snapshotImage = File.ReadAllBytes(snapshotPath);
            }
            else
            {
                // Fallback to an empty image if the file is missing
                snapshotImage = new byte[0];
                Console.WriteLine($"Warning: '{snapshotPath}' not found. Using empty image data.");
            }

            // Prepare embedded OLE data (e.g., an Excel file to embed)
            byte[] embeddedData;
            const string embeddedPath = "embedded.xlsx";
            if (File.Exists(embeddedPath))
            {
                embeddedData = File.ReadAllBytes(embeddedPath);
            }
            else
            {
                // Fallback to an empty byte array if the file is missing
                embeddedData = new byte[0];
                Console.WriteLine($"Warning: '{embeddedPath}' not found. Using empty embedded data.");
            }

            // Add the original OLE object to the source worksheet
            int originalIndex = sourceSheet.OleObjects.Add(2, 2, 200, 200, snapshotImage);
            OleObject originalOle = sourceSheet.OleObjects[originalIndex];

            // Set the embedded object data and its format
            originalOle.ObjectData = embeddedData;
            originalOle.FileFormatType = FileFormatType.Xlsx;

            // Add a new worksheet that will receive the cloned OLE object
            Worksheet targetSheet = workbook.Worksheets.Add("Clone");

            // Clone the OLE object by adding a new one with the same visual data
            int cloneIndex = targetSheet.OleObjects.Add(
                originalOle.UpperLeftRow,
                originalOle.UpperLeftColumn,
                originalOle.Height,
                originalOle.Width,
                originalOle.ImageData);

            OleObject clonedOle = targetSheet.OleObjects[cloneIndex];

            // Copy the embedded data and relevant properties from the original
            clonedOle.ObjectData = originalOle.ObjectData;
            clonedOle.FileFormatType = originalOle.FileFormatType;
            clonedOle.DisplayAsIcon = originalOle.DisplayAsIcon;
            clonedOle.ProgID = originalOle.ProgID;

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ClonedOleObject.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved as 'ClonedOleObject.xlsx'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
