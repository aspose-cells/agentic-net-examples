// Title: How to set a custom display label for an OLE object in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom label for embedded OLE object in Excel | change OLE object caption programmatically with Aspose.Cells .NET | assign display name to OLE object added to worksheet using Aspose.Cells | set OleObject.Label property after adding DOCX as OLE in C# | update OLE object label in existing workbook with Aspose.Cells API
// Tags: Aspose.Cells OLE object label assignment | C# OleObject.Label usage | embed DOCX as OLE with Aspose.Cells | customize OLE object caption in Excel | Aspose.Cells worksheet OLE manipulation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, reads a DOCX file into a byte array, adds it as an OLE object to the first worksheet, assigns a custom display label via the OleObject.Label property, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for the workbook and the OLE source file
            string workbookPath = "input.xlsx";
            string oleSourcePath = @"C:\Docs\sample.docx";
            string outputPath = "output.xlsx";

            // Verify that the required files exist
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            if (!File.Exists(oleSourcePath))
            {
                Console.WriteLine($"OLE source file not found: {oleSourcePath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(workbookPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Read the OLE object data as a byte array
            byte[] oleData = File.ReadAllBytes(oleSourcePath);

            // Add the OLE object to the sheet (upper left row, column, lower right row, column, data)
            int oleIndex = sheet.OleObjects.Add(1, 1, 10, 5, oleData);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Assign a custom display label to the OLE object
            ole.Label = "Custom Document";

            // Save the workbook with the updated OLE object
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
