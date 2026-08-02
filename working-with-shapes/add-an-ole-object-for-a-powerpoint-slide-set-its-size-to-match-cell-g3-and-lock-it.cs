// Title: Lock an OLE object in cell G3 with matching size using Aspose.Cells for .NET
// Description: Shows how to obtain the pixel height and width of cell G3, embed a PNG (or any file) as an OLE object at that location, size the object to the cell, lock it, and save the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells OLE object | C# add OLE to Excel | lock OLE object Aspose | set OLE size to cell | embed PowerPoint slide as OLE | Excel cell dimensions pixel | Aspose.Cells worksheet protection | C# Excel OLE embed
// Common Searches: Aspose.Cells add OLE object to specific cell | C# lock OLE object in Excel worksheet | size OLE object to cell dimensions Aspose | embed PowerPoint slide as OLE in Excel using Aspose | set OLE object width height from cell G3
// Developer Intent: Add an OLE object to cell G3, match its width and height to the cell, and lock it so it cannot be edited.
// Use Cases: Create a protected template where embedded charts are read‑only. | Insert a PowerPoint slide preview that aligns with a designated cell in a report. | Place a placeholder image that automatically fits the cell for later replacement. | Lock embedded objects to prevent changes when the worksheet is protected.
// AI Prompts: Write C# code with Aspose.Cells to insert a locked OLE object into cell G3, using the cell's pixel dimensions for size. | Show how to retrieve G3 row height and column width in pixels and apply them to an OLE object's Height and Width properties in Aspose.Cells. | Explain how to lock an OLE object and protect the worksheet so the object remains uneditable. | Provide an example of embedding a PowerPoint slide as an OLE object in an Excel file with Aspose.Cells and locking it.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to obtain the pixel height and width of cell G3, embed a PNG (or any file) as an OLE object at that location, size the object to the cell, lock it, and save the workbook with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Cell G3 (zero‑based indices: row 2, column 6)
            int rowIndex = 2;
            int colIndex = 6;

            // Get the pixel dimensions of the cell
            double cellHeightPx = sheet.Cells.GetRowHeightPixel(rowIndex);
            double cellWidthPx = sheet.Cells.GetColumnWidthPixel(colIndex);
            int height = (int)Math.Round(cellHeightPx);
            int width = (int)Math.Round(cellWidthPx);

            // 1x1 transparent PNG (base64 encoded)
            const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
            byte[] imageData = Convert.FromBase64String(pngBase64);

            // Add the OLE object at G3 with size matching the cell
            int oleIndex = sheet.OleObjects.Add(rowIndex, colIndex, height, width, imageData);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Lock the OLE object so it cannot be modified when the sheet is protected
            ole.IsLocked = true;

            // Save the workbook
            string outputPath = "OleObjectInPowerPointSlide.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
