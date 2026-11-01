// Title: Assign a GUID‑based unique Name to an OLE object after adding it to an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that inserts a Word file as an OLE object into cell A1 of a worksheet using Aspose.Cells, then assigns a GUID‑generated string to the OleObject.Name property. | Create a method that adds an OLE object to a worksheet and guarantees a distinct Name value by concatenating a prefix with Guid.NewGuid().ToString("N"). | Refactor an existing Aspose.Cells example so that every newly added OLE object receives a unique identifier via its Name property.
// Common Searches: asp.net c# how to set a unique name for an ole object in an Aspose.Cells workbook | set OleObject.Name to a GUID after inserting an OLE object with Aspose.Cells for .NET | example code for embedding a Word document as an OLE object and naming it uniquely in Excel using Aspose.Cells | c# Aspose.Cells best practice for assigning distinct identifiers to OLE objects
// Tags: Aspose.Cells add OLE object C# | OleObject.Name unique identifier | GUID naming for OLE objects | embed Word document as OLE in Excel | set OLE object name property Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, embeds a Word document as an OLE object at cell A1, generates a GUID‑based string, assigns it to the OleObject.Name property to ensure uniqueness, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the file that will be embedded as an OLE object
            string oleSourcePath = "sample.docx";

            // Ensure the source file exists
            if (!File.Exists(oleSourcePath))
            {
                Console.WriteLine($"Source file not found: {oleSourcePath}");
                return;
            }

            // Read the file into a byte array (required by the Add method)
            byte[] oleData = File.ReadAllBytes(oleSourcePath);

            // Insert the OLE object at cell A1 (row 0, column 0) with width and height of 100 points
            // Add method returns the index of the newly added OLE object
            int oleIndex = sheet.OleObjects.Add(0, 0, 100, 100, oleData);

            // Retrieve the OleObject instance using the returned index
            OleObject ole = sheet.OleObjects[oleIndex];

            // Assign a unique identifier to the OLE object's Name property
            ole.Name = "OleObject_" + Guid.NewGuid().ToString("N");

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
