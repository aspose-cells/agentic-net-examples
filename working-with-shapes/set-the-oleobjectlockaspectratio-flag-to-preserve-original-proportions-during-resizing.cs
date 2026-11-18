// Title: Preserving OLE Object Aspect Ratio When Resizing with Aspose.Cells for .NET (LockAspectRatio Not Supported)
// Common Searches: Aspose.Cells C# keep OLE object proportions during resize | How to lock aspect ratio of an OLE object in an Excel file using Aspose.Cells | Is there a LockAspectRatio property for OleObject in Aspose.Cells .NET | Resize OLE object without distortion in Aspose.Cells workbook | Alternative to LockAspectRatio for OleObject in Aspose.Cells C# | Set IsObjectLocked for OLE objects in Aspose.Cells
// Tags: OleObject resizing preserve aspect ratio Aspose.Cells | Aspose.Cells OLE object scaling limitation | C# Aspose.Cells lock OLE object editing | Excel workbook OLE object aspect ratio handling | Aspose.Cells IsObjectLocked property usage | Aspose.Cells .NET OLE object manipulation

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook, checks for OLE objects on the first worksheet, accesses the first OleObject, and demonstrates that Aspose.Cells does not provide a LockAspectRatio property. It notes that only the IsObjectLocked flag can be used to prevent editing, and then saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify the desired one)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one OleObject on the worksheet
            if (sheet.OleObjects.Count > 0)
            {
                // Access the first OleObject on the worksheet
                OleObject oleObject = sheet.OleObjects[0];

                // Note: Aspose.Cells OleObject does not expose a LockAspectRatio property.
                // If needed, you can lock the object to prevent editing using IsObjectLocked.
                // oleObject.IsObjectLocked = true; // Uncomment if locking is required and supported.
            }
            else
            {
                Console.WriteLine("No OleObjects found on the first worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
