// Title: Insert and lock a Visio diagram as a cell‑sized OLE object with Aspose.Cells for .NET
// Description: Creates a new workbook, calculates the pixel width and height of a target cell, adds an OLE object sized to that cell, embeds a .vsdx Visio file as a non‑icon object, locks it, sets PlacementType.MoveAndSize, and saves the file.
// Keywords: Aspose.Cells | C# | OLE object | Visio embed | lock OLE | cell dimensions | PlacementType.MoveAndSize | Excel automation | .vsdx | embed diagram
// Common Searches: embed Visio file in Excel using Aspose.Cells | lock OLE object in Aspose.Cells workbook | size OLE object to match Excel cell | PlacementType.MoveAndSize example C# | Aspose.Cells add OLE object from byte array
// Developer Intent: Embed a Visio diagram as a locked OLE object that automatically fits and moves with a specific worksheet cell.
// Use Cases: Generate reports where each row displays a Visio diagram locked inside its cell. | Create a template with pre‑locked Visio OLE objects that resize with underlying cells. | Automate bulk insertion of Visio diagrams into different cells while preventing user edits.
// AI Prompts: Write C# code with Aspose.Cells to embed a Visio .vsdx file as a non‑icon OLE object at cell D10, size it to the cell's pixel dimensions, lock it, and set it to move and resize with the cell. | Explain how to obtain a worksheet cell's pixel width and height in Aspose.Cells and use those values for OLE object sizing. | Provide step‑by‑step instructions to embed a Visio diagram as a locked OLE object and ensure it stays protected when the sheet is protected.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Creates a new workbook, calculates the pixel width and height of a target cell, adds an OLE object sized to that cell, embeds a .vsdx Visio file as a non‑icon object, locks it, sets PlacementType.MoveAndSize, and saves the file.
class InsertVisioOleObject
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Target cell where the OLE object will be placed (zero‑based indexes)
            int targetRow = 5;      // Row 6 in Excel UI
            int targetColumn = 2;   // Column C in Excel UI

            // Determine the pixel size of the target cell
            int widthPx = sheet.Cells.GetColumnWidthPixel(targetColumn);
            int heightPx = sheet.Cells.GetRowHeightPixel(targetRow);

            // Add an OLE object at the target cell with the calculated size.
            // An empty byte array is supplied for the image data (required by the API).
            byte[] placeholderImage = new byte[0];
            int oleIndex = sheet.OleObjects.Add(targetRow, targetColumn, heightPx, widthPx, placeholderImage);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Load the Visio diagram to be embedded
            string visioPath = "diagram.vsdx";
            if (!File.Exists(visioPath))
                throw new FileNotFoundException($"Visio file not found: {visioPath}");

            byte[] visioData = File.ReadAllBytes(visioPath);

            // Embed the Visio file (not linked, not displayed as an icon)
            ole.SetEmbeddedObject(
                linkToFile: false,
                objectData: visioData,
                sourceFileName: Path.GetFileName(visioPath),
                displayAsIcon: false,
                label: string.Empty);

            // Lock the OLE object so it cannot be modified when the sheet is protected
            ole.IsLocked = true;

            // Ensure the object moves and resizes with its underlying cells
            ole.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            string outputPath = "VisioOleObject.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
