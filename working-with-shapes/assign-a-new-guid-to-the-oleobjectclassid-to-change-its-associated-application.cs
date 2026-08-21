// Title: Assign a New GUID to OleObject.ClassIdentifier with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add an OLE object, generate a fresh GUID, assign its byte array to OleObject.ClassIdentifier, optionally set a ProgID (e.g., "Excel.Sheet.12"), save the file, reload it, and verify that the GUID is persisted.
// Keywords: Aspose.Cells | C# | OleObject | ClassIdentifier | GUID | ProgID | Excel OLE object | generate GUID | embed OLE | change OLE application | persist OLE object
// Common Searches: Aspose.Cells assign new GUID to OleObject | set OleObject ClassIdentifier C# | change OLE object ProgID with Aspose.Cells | verify GUID of embedded OLE object after save | how to generate and apply GUID to Excel OLE object
// Developer Intent: Programmatically replace the default ClassIdentifier of an OleObject with a newly generated GUID to associate it with a different application.
// Use Cases: Embed an OLE object in an Excel workbook and bind it to a custom application via a unique GUID. | Update the ProgID of an existing OLE object after assigning a new ClassIdentifier. | Save and reload a workbook to confirm that the custom GUID is correctly stored in the file.
// AI Prompts: Write C# code using Aspose.Cells to add an OLE object, assign a newly created GUID to its ClassIdentifier, set a custom ProgID, and validate the GUID after saving the workbook. | Explain the relationship between the ClassIdentifier byte array and a GUID, and how Aspose.Cells encodes this information in an Excel file. | Suggest robust error‑handling patterns for GUID assignment to OleObject.ClassIdentifier in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add an OLE object, generate a fresh GUID, assign its byte array to OleObject.ClassIdentifier, optionally set a ProgID (e.g., "Excel.Sheet.12"), save the file, reload it, and verify that the GUID is persisted.
    public class OleObjectAssignNewGuidDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add an OLE object (empty data for demonstration)
                int oleIndex = sheet.OleObjects.Add(5, 5, 150, 100, new byte[0]);
                OleObject oleObject = sheet.OleObjects[oleIndex];

                // Generate a new GUID and assign its byte representation to ClassIdentifier
                Guid newGuid = Guid.NewGuid();
                oleObject.ClassIdentifier = newGuid.ToByteArray();

                // Optionally set ProgID to reflect the target application (e.g., Excel)
                oleObject.ProgID = "Excel.Sheet.12";

                // Save the workbook
                string outputPath = "OleObjectWithNewGuid.xlsx";
                workbook.Save(outputPath);

                // Verify the file exists before loading
                if (!File.Exists(outputPath))
                {
                    Console.WriteLine($"Error: The file '{outputPath}' was not found after saving.");
                    return;
                }

                // Load the workbook to verify the GUID was persisted
                Workbook loadedWorkbook = new Workbook(outputPath);
                OleObject loadedOle = loadedWorkbook.Worksheets[0].OleObjects[0];
                Guid loadedGuid = new Guid(loadedOle.ClassIdentifier);

                Console.WriteLine("Original GUID : " + newGuid);
                Console.WriteLine("Loaded GUID   : " + loadedGuid);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            OleObjectAssignNewGuidDemo.Run();
        }
    }
}
