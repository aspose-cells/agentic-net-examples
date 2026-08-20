// Title: Audit CLSID (ClassIdentifier) of embedded OLE objects in Excel using Aspose.Cells for .NET
// Description: Loads a workbook, iterates every worksheet and its OleObjects collection, extracts the 16‑byte ClassIdentifier, converts it to a GUID, and prints the sheet name, OLE index and CLSID. Ideal for compliance checks and object validation.
// Keywords: Aspose.Cells OLE CLSID | C# read OleObject ClassIdentifier | Excel embedded OLE GUID | Aspose.Cells audit OLE objects | .NET extract OLE ClassIdentifier | GitHub Aspose.Cells OLE example | US developers Aspose.Cells | Europe .NET Excel OLE
// Common Searches: how to get CLSID of an OLE object with Aspose.Cells | Aspose.Cells retrieve OleObject ClassIdentifier C# | convert OLE ClassIdentifier to GUID in .NET | list embedded OLE objects in Excel workbook Aspose | audit Excel OLE objects for compliance
// Developer Intent: Obtain the CLSID of each embedded OLE object in an Excel file for verification or reporting.
// Use Cases: Cross‑check OLE objects against an approved CLSID whitelist. | Produce a compliance report that lists worksheet names, OLE positions and their GUIDs. | Detect missing or corrupted OLE entries by flagging invalid ClassIdentifier arrays.
// AI Prompts: Generate C# code with Aspose.Cells that enumerates all OLE objects in a workbook, extracts their CLSID GUIDs, and writes the results to a CSV file. | Create a method that receives a file path and returns a dictionary mapping worksheet names to collections of OLE CLSIDs, handling null or malformed identifiers gracefully. | Explain how to compare extracted CLSID GUIDs with a predefined whitelist and highlight any non‑compliant OLE objects.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, iterates every worksheet and its OleObjects collection, extracts the 16‑byte ClassIdentifier, converts it to a GUID, and prints the sheet name, OLE index and CLSID. Ideal for compliance checks and object validation.
class OleObjectClassIdAudit
{
    static void Main()
    {
        // Load an existing workbook that contains OLE objects
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through each OLE object in the worksheet
            for (int i = 0; i < sheet.OleObjects.Count; i++)
            {
                OleObject ole = sheet.OleObjects[i];

                // Retrieve the ClassIdentifier (CLSID) byte array
                byte[] classIdBytes = ole.ClassIdentifier;

                // Convert the byte array to a GUID string if it has the expected length (16 bytes)
                string classIdGuid = (classIdBytes != null && classIdBytes.Length == 16)
                    ? new Guid(classIdBytes).ToString()
                    : "Invalid or missing ClassIdentifier";

                // Output the auditing information
                Console.WriteLine($"Worksheet: {sheet.Name}, OLE Index: {i}, ClassIdentifier (GUID): {classIdGuid}");
            }
        }
    }
}
