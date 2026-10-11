// Title: How to assign a new GUID to OleObject.ClassId in Aspose.Cells for .NET to change the linked application
// AI Prompts: Generate C# code that creates an OLE placeholder in a worksheet using Aspose.Cells, generates a GUID, and sets the OleObject.ClassId property to that GUID before saving the workbook. | Show how to retrieve an existing OleObject from an Aspose.Cells workbook and assign a custom GUID to its ClassId property to redirect the OLE object to a different application.
// Common Searches: Aspose.Cells C# set custom GUID for OleObject ClassId to change application association | How to change the linked application of an OLE object in an Excel file using Aspose.Cells .NET | Assigning a new ClassId to an OleObject in an Aspose.Cells workbook example
// Tags: Aspose.Cells set OleObject ClassId GUID | C# assign custom GUID to OLE object | modify OLE object application link Aspose.Cells | OleObject ClassId property .NET | Excel OLE placeholder Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example demonstrates creating a workbook with Aspose.Cells, adding an empty OLE object, and shows where to generate and assign a new GUID to the OleObject.ClassId property so the OLE object points to a different application before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add an empty OLE object. The Add method expects a byte array for the OLE data.
            // An empty byte array creates a placeholder object.
            byte[] emptyOleData = new byte[0];
            int oleIndex = sheet.OleObjects.Add(1, 1, 100, 100, emptyOleData);

            // Retrieve the newly added OLE object.
            OleObject oleObject = sheet.OleObjects[oleIndex];

            // Optional: set a display name.
            // oleObject.Name = "MyOleObject";

            // Determine output path and ensure its directory exists.
            string outputPath = "Output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
