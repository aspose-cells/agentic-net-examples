// Title: C# – Insert OLE Object and Freeze Panes with Aspose.Cells for .NET
// Description: Demonstrates how to add an OLE object (with a preview image or placeholder) to a worksheet, set its ProgID, display‑as‑icon label, and freeze rows/columns up to the object's top‑left cell, then save the workbook as XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | OLE object | embed Excel file | preview image | display as icon | ProgID | freeze panes | freeze rows | freeze columns | GitHub example | coding‑agent | worksheet layout
// Common Searches: How to embed an Excel file as an OLE object with Aspose.Cells C# | Aspose.Cells freeze panes around an OLE object | C# code to add OLE object and freeze rows in a workbook | Insert OLE object with preview image using Aspose.Cells for .NET | Set ProgID and display‑as‑icon for OLE objects in Aspose.Cells
// Developer Intent: Add an OLE object to a worksheet and freeze the surrounding rows/columns for a fixed layout.
// Use Cases: Create a report that embeds a secondary workbook as an icon while keeping header rows visible. | Place an OLE chart on a dashboard sheet and lock surrounding data rows so the chart remains in view. | Generate a template with an attached spreadsheet as an OLE object and freeze panes to lock the view to the top of the embedded file.
// AI Prompts: Write C# code using Aspose.Cells to insert an OLE object with a custom preview image and freeze panes at the object's position. | Explain how to configure ProgID, display‑as‑icon, and label for an OLE object in Aspose.Cells. | Provide robust error handling for embedding files as OLE objects with Aspose.Cells in a .NET application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add an OLE object (with a preview image or placeholder) to a worksheet, set its ProgID, display‑as‑icon label, and freeze rows/columns up to the object's top‑left cell, then save the workbook as XLSX using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Paths to the preview image and the file to embed (replace with actual paths)
            string imagePath = "preview.png";
            string embeddedFilePath = "sample.xlsx";

            // Load image bytes; if the file does not exist, use a 1x1 white PNG placeholder
            byte[] imageData;
            if (File.Exists(imagePath))
            {
                imageData = File.ReadAllBytes(imagePath);
            }
            else
            {
                // Base64-encoded 1x1 white PNG
                const string placeholderBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
                imageData = Convert.FromBase64String(placeholderBase64);
            }

            // Load the file that will be embedded inside the OLE object
            byte[] objectData = Array.Empty<byte>();
            if (File.Exists(embeddedFilePath))
            {
                objectData = File.ReadAllBytes(embeddedFilePath);
            }
            else
            {
                Console.WriteLine($"Embedded file not found: {embeddedFilePath}");
            }

            // Define the position (top‑left cell) and size (in pixels) of the OLE object
            int topRow = 5;        // zero‑based row index
            int leftColumn = 2;    // zero‑based column index
            int height = 200;      // pixel height
            int width = 300;       // pixel width

            // Add the OLE object to the worksheet
            int oleIndex = sheet.OleObjects.Add(topRow, leftColumn, height, width, imageData);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Set the embedded file data and display properties
            ole.ObjectData = objectData;
            ole.ProgID = "Excel.Sheet.8";               // ProgID for an embedded Excel workbook
            ole.DisplayAsIcon = true;                  // Show as an icon
            ole.Label = Path.GetFileName(embeddedFilePath); // Icon label

            // Freeze rows and columns up to the OLE object's top‑left cell
            sheet.FreezePanes(topRow, leftColumn, topRow, leftColumn);

            // Save the workbook
            workbook.Save("OleObjectWithFreeze.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
