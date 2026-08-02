// Title: Assign a New GUID to an OleObject’s ClassIdentifier Using Aspose.Cells for .NET (C#)
// Description: C# sample that creates an Excel workbook with Aspose.Cells, inserts an OLE object, generates a fresh GUID, assigns it to OleObject.ClassIdentifier, optionally sets the ProgID (e.g., Word.Document.12), saves the file, reloads it, and verifies that the GUID and ProgID persist.
// Keywords: Aspose.Cells | C# | OleObject | ClassIdentifier | GUID | ProgID | Excel OLE object | set OLE GUID | change OLE application | generate GUID for OLE | persist OLE object
// Common Searches: Aspose.Cells assign GUID to OleObject | C# change OleObject ClassIdentifier | set ProgID for Excel OLE object with Aspose.Cells | verify OLE GUID after saving workbook | generate and persist custom GUID for OLE in Excel | how to update OleObject application association
// Developer Intent: Update an OleObject’s ClassIdentifier with a new GUID (and optional ProgID) to alter its linked application via Aspose.Cells.
// Use Cases: Create a workbook, add an OLE placeholder, generate a GUID, assign it to ClassIdentifier, set a ProgID, and save the file. | Reload the saved workbook to read back the OleObject, convert the stored ClassIdentifier bytes to a Guid, and confirm it matches the original. | Replace the default OLE application by assigning a custom GUID and updating the ProgID to point to a different program.
// AI Prompts: Write C# code using Aspose.Cells that adds an OLE object, generates a new GUID, assigns it to ClassIdentifier, sets ProgID to "Word.Document.12", saves the workbook, and validates the GUID after loading. | Explain the steps to persist a custom GUID for an OleObject in an Excel file with Aspose.Cells and retrieve it later in C#. | Provide a step‑by‑step tutorial for changing the associated application of an OLE object by updating its ClassIdentifier and ProgID using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // C# sample that creates an Excel workbook with Aspose.Cells, inserts an OLE object, generates a fresh GUID, assigns it to OleObject.ClassIdentifier, optionally sets the ProgID (e.g., Word.Document.12), saves the file, reloads it, and verifies that the GUID and ProgID persist.
    public class OleObjectAssignNewGuidDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add an OLE object with placeholder data
                int oleIndex = sheet.OleObjects.Add(10, 10, 200, 100, new byte[0]);
                OleObject oleObject = sheet.OleObjects[oleIndex];

                // Generate a new GUID and assign it to the ClassIdentifier
                Guid newGuid = Guid.NewGuid();
                oleObject.ClassIdentifier = newGuid.ToByteArray();

                // Optionally set ProgID
                oleObject.ProgID = "Word.Document.12";

                // Save the workbook
                string filePath = "OleObjectWithNewGuid.xlsx";
                workbook.Save(filePath);

                // Verify the GUID was persisted
                if (File.Exists(filePath))
                {
                    Workbook loadedWorkbook = new Workbook(filePath);
                    OleObject loadedOle = loadedWorkbook.Worksheets[0].OleObjects[0];
                    Guid loadedGuid = new Guid(loadedOle.ClassIdentifier);
                    Console.WriteLine("Original GUID: " + newGuid);
                    Console.WriteLine("Loaded GUID  : " + loadedGuid);
                    Console.WriteLine("ProgID       : " + loadedOle.ProgID);
                }
                else
                {
                    Console.WriteLine("File not found: " + filePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            OleObjectAssignNewGuidDemo.Run();
        }
    }
}
