// Title: Insert a PDF as an OLE object into cell H4 of an Excel worksheet and define its height and width using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that reads a PDF file into a byte array and embeds it as an OLE object at cell H4, setting the object height to 200 points and width to 300 points. | Show how to add a PDF OLE object to a worksheet using the OleObjects.Add method, then save the workbook with Aspose.Cells. | Adapt the example to display the inserted PDF OLE object as an icon and assign a custom icon caption.
// Common Searches: Aspose.Cells C# embed PDF as OLE object in a specific worksheet cell | How to set height and width of an OLE object added with Aspose.Cells | Add PDF OLE object to Excel at cell H4 using Aspose.Cells .NET
// Tags: OleObjects.Add PDF embedding Aspose.Cells | set OLE object dimensions Excel .NET | embed PDF as OLE object worksheet cell | read PDF to byte array Aspose.Cells | display OLE object as icon Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, loads a PDF file into a byte array, inserts it as an OLE object positioned at cell H4 with a height of 200 points and a width of 300 points, disables icon display, and saves the result as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Initialize a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the PDF file that will be embedded as an OLE object
            string pdfPath = "sample.pdf";

            // Verify that the PDF file exists to avoid FileNotFoundException
            if (!File.Exists(pdfPath))
            {
                Console.WriteLine($"Error: The file '{pdfPath}' was not found.");
                return;
            }

            // Read the PDF file into a byte array (required by OleObjects.Add overload)
            byte[] pdfData = File.ReadAllBytes(pdfPath);

            // Define the target cell (H4) where the OLE object will be placed
            // Row and column indices are zero‑based: H = column 7, 4th row = index 3
            int targetRow = 3;      // Row 4
            int targetColumn = 7;   // Column H

            // Desired size of the OLE object (in points)
            int objectHeight = 200; // Height
            int objectWidth = 300;  // Width

            // Add the OLE object to the worksheet using the byte array.
            // The Add method returns the index of the newly added OleObject.
            int oleIndex = sheet.OleObjects.Add(targetRow, targetColumn, objectHeight, objectWidth, pdfData);

            // Retrieve the OleObject instance using the returned index
            OleObject ole = sheet.OleObjects[oleIndex];

            // Optional: set additional properties (e.g., display as icon)
            ole.DisplayAsIcon = false;

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
