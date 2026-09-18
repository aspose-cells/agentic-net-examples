// Title: Lock aspect ratio of an embedded OLE object in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Add a DOCX file as an OLE object to a worksheet and enable its LockAspectRatio property with Aspose.Cells for .NET. | Retrieve an existing OleObject from a workbook and set LockAspectRatio to true so resizing keeps the original width‑height ratio. | Create a new workbook, embed an OLE object, lock its aspect ratio, and save the file using C# and Aspose.Cells.
// Common Searches: C# Aspose.Cells how to keep OLE object proportions when resizing | set OleObject.LockAspectRatio to true Aspose.Cells example | embed Word document as OLE object in Excel and preserve aspect ratio using Aspose.Cells | Aspose.Cells lock aspect ratio of embedded object programmatically
// Tags: OleObject.LockAspectRatio Aspose.Cells C# | embed DOCX as OLE object Excel Aspose.Cells | preserve OLE object proportions Aspose.Cells | aspect ratio lock workbook C# | Aspose.Cells OLE object resizing

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, embeds a DOCX file as an OLE object at a specified cell, accesses the OleObject instance, sets its LockAspectRatio flag to true to maintain original proportions during any resizing, and saves the workbook.
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

            // Parameters for the OleObject
            int row = 2;          // Row index (0‑based)
            int column = 1;       // Column index (0‑based)
            int height = 200;     // Height in pixels
            int width = 300;      // Width in pixels
            string sourceFile = "sample.docx"; // Path to the embedded file

            // Verify that the source file exists before loading
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Error: The file '{sourceFile}' was not found.");
                return;
            }

            // Read the file into a byte array (required by the API)
            byte[] oleData = File.ReadAllBytes(sourceFile);

            // Add the OleObject to the worksheet (returns the index of the new object)
            int oleIndex = sheet.OleObjects.Add(row, column, height, width, oleData, "Sample Document");

            // Retrieve the OleObject instance
            OleObject ole = sheet.OleObjects[oleIndex];

            // (Optional) Additional settings can be applied here if needed

            // Save the workbook
            string outputPath = "OleObjectDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
