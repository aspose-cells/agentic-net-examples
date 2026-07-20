// Title: Read OLE Object CLSID (ClassIdentifier) in Excel with Aspose.Cells for .NET
// Description: Loads a workbook, iterates each worksheet’s OleObjects collection, extracts the ClassIdentifier byte array, converts it to a readable hex string, and prints the worksheet name, OLE index, ProgID, and CLSID. The sample also shows how to create a new workbook, add an OLE object with a custom 16‑byte CLSID, set its ProgID, and save the file for testing.
// Keywords: Aspose.Cells OLE CLSID | ClassIdentifier property .NET | read embedded OLE object GUID | audit Excel OLE objects | convert byte[] to hex Aspose | C# OleObject ClassIdentifier | Excel security scan OLE
// Common Searches: how to get CLSID of OLE object using Aspose.Cells | Aspose.Cells read ClassIdentifier from Excel | C# list embedded OLE objects in workbook | convert OLE ClassIdentifier to string | Aspose.Cells audit embedded OLE content
// Developer Intent: Retrieve the ClassIdentifier (CLSID) of each embedded OLE object in an Excel workbook for auditing or compliance verification.
// Use Cases: Generate a compliance report that lists worksheet, OLE index, ProgID, and CLSID for all embedded objects. | Detect prohibited or unknown OLE content by scanning CLSIDs across workbooks. | Create a test workbook with a known CLSID to validate that the ClassIdentifier persists after saving.
// AI Prompts: Write C# code with Aspose.Cells to enumerate OLE objects in an Excel file and output their CLSID values. | Show how to assign a custom 16‑byte ClassIdentifier to a new OleObject and save the workbook. | Explain how to format the byte[] ClassIdentifier as a standard GUID string in C#.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, iterates each worksheet’s OleObjects collection, extracts the ClassIdentifier byte array, converts it to a readable hex string, and prints the worksheet name, OLE index, ProgID, and CLSID. The sample also shows how to create a new workbook, add an OLE object with a custom 16‑byte CLSID, set its ProgID, and save the file for testing.
class OleObjectClassIdAudit
{
    static void Main()
    {
        // Path to the workbook that contains OLE objects to audit
        string inputPath = "SampleWithOle.xlsx";

        Workbook workbook = null;

        // Load the workbook only if the file exists
        if (File.Exists(inputPath))
        {
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine($"Input file '{inputPath}' not found. Skipping audit.");
        }

        // Perform audit if workbook was loaded successfully
        if (workbook != null)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                for (int i = 0; i < sheet.OleObjects.Count; i++)
                {
                    OleObject ole = sheet.OleObjects[i];

                    // Retrieve the ClassIdentifier (CLSID) as a byte array
                    byte[] clsid = ole.ClassIdentifier;

                    // Convert the CLSID to a readable hex string (or indicate if none)
                    string clsidHex = (clsid != null && clsid.Length > 0)
                        ? BitConverter.ToString(clsid).Replace("-", "")
                        : "None";

                    // Output audit information
                    Console.WriteLine($"Worksheet: {sheet.Name}, OLE Index: {i}");
                    Console.WriteLine($"ProgID: {ole.ProgID}");
                    Console.WriteLine($"ClassIdentifier (CLSID): {clsidHex}");
                    Console.WriteLine();
                }
            }
        }

        // ------------------------------------------------------------
        // Demonstration: create a new workbook and add an OLE object
        // with a known ClassIdentifier for testing purposes (create rule)
        // ------------------------------------------------------------
        try
        {
            Workbook newWb = new Workbook();
            Worksheet newSheet = newWb.Worksheets[0];

            // Sample CLSID (16-byte array)
            byte[] sampleClsId = new byte[]
            {
                0x01,0x02,0x03,0x04,
                0x05,0x06,0x07,0x08,
                0x09,0x0A,0x0B,0x0C,
                0x0D,0x0E,0x0F,0x10
            };

            // Add an empty OLE object and set its properties
            int oleIdx = newSheet.OleObjects.Add(5, 5, 100, 50, new byte[0]);
            OleObject newOle = newSheet.OleObjects[oleIdx];
            newOle.ClassIdentifier = sampleClsId;
            newOle.ProgID = "Excel.Sheet.12";

            // Save the new workbook (save rule)
            string outputPath = "AuditDemo.xlsx";
            newWb.Save(outputPath);
            Console.WriteLine($"Demo workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during demo workbook creation: {ex.Message}");
        }
    }
}
