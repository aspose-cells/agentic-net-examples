// Title: How to set a custom display label for an embedded OLE object in Excel using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set OleObject label after embedding a Word file | how to change the display name of an OLE object in an Excel workbook with Aspose.Cells | Aspose.Cells .NET assign custom label to embedded OLE object example | set OleObject.Label property in Aspose.Cells C# code | embed DOCX as OLE in Excel and customize its label using Aspose.Cells
// Tags: Aspose.Cells set OleObject.Label C# | embed Word document as OLE in Excel using Aspose.Cells | custom display label for Excel OLE object | Aspose.Cells OLE object properties manipulation | C# add OLE object to worksheet with label

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, embeds a Word document (sample.docx) as an OLE object within a defined cell range, assigns a custom display label via the OleObject.Label property, and saves the file as Output.xlsx.
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

            // Define the cell range where the OLE object will be placed (zero‑based indices)
            int upperLeftRow = 2;
            int upperLeftColumn = 1;
            int lowerRightRow = 10;
            int lowerRightColumn = 5;

            // Path to the file that will be embedded as an OLE object
            string oleFilePath = "sample.docx";

            // Ensure the OLE source file exists
            if (!File.Exists(oleFilePath))
                throw new FileNotFoundException($"The OLE source file '{oleFilePath}' was not found.");

            // Read the file into a byte array (required by the Add method)
            byte[] oleData = File.ReadAllBytes(oleFilePath);

            // Approximate height and width in pixels based on the cell range
            // (Aspose.Cells expects height and width, not lower‑right cell indices)
            int height = (lowerRightRow - upperLeftRow + 1) * 15;   // 15 points per row (approx.)
            int width = (lowerRightColumn - upperLeftColumn + 1) * 64; // 64 points per column (approx.)

            // Add the OLE object to the worksheet.
            // The Add method returns the index of the newly added object.
            int oleIndex = sheet.OleObjects.Add(
                upperLeftRow,
                upperLeftColumn,
                height,
                width,
                oleData,
                Path.GetExtension(oleFilePath));

            // Retrieve the OleObject instance using the returned index
            OleObject ole = sheet.OleObjects[oleIndex];

            // Assign a custom display label to the OLE object
            ole.Label = "My Custom Document";

            // Save the workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
