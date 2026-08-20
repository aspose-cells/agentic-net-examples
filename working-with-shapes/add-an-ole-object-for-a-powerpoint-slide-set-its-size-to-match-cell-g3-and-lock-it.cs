// Title: Add a PowerPoint OLE object to cell G3, size it to the cell and lock it – Aspose.Cells C#
// Description: C# example that creates a workbook, reads the pixel width and height of cell G3, inserts a PowerPoint OLE object with matching dimensions, locks the object to prevent editing, and saves the file as OleObjectForPowerPoint.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells OLE object | PowerPoint OLE Excel | C# resize OLE to cell | lock OLE object Aspose | cell pixel dimensions Aspose.Cells | embed PowerPoint slide Excel | Aspose.Cells add OLE example
// Common Searches: how to embed a PowerPoint slide as OLE in Excel with Aspose.Cells | resize OLE object to match Excel cell size C# | lock OLE object after insertion Aspose.Cells | get cell pixel width and height Aspose.Cells | add OLE object to specific cell Aspose.Cells .NET
// Developer Intent: Insert a PowerPoint OLE object into cell G3, size it to the cell’s pixel dimensions, and lock the object using Aspose.Cells for .NET.
// Use Cases: Embedding a PowerPoint slide in a financial report so it aligns perfectly with a designated cell. | Automating layout consistency by programmatically matching OLE object size to target cells. | Preventing users from moving or editing embedded PowerPoint content by locking the OLE object.
// AI Prompts: Write C# code with Aspose.Cells to add a PowerPoint OLE object to cell G3, apply the exact pixel width and height of the cell, and lock the object. | Show how to obtain a cell’s pixel dimensions in Aspose.Cells and use them to size an OLE object inserted into the worksheet. | Explain how to replace the placeholder PNG with a real PowerPoint file stream while keeping the OLE object locked in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, reads the pixel width and height of cell G3, inserts a PowerPoint OLE object with matching dimensions, locks the object to prevent editing, and saves the file as OleObjectForPowerPoint.xlsx using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Cell G3 (zero‑based indices: row 2, column 6)
            int rowIndex = 2;
            int columnIndex = 6;

            // Get the pixel dimensions of the cell
            double cellHeightPx = worksheet.Cells.GetRowHeightPixel(rowIndex);
            double cellWidthPx = worksheet.Cells.GetColumnWidthPixel(columnIndex);

            // Convert to integer pixel values
            int height = (int)Math.Round(cellHeightPx);
            int width = (int)Math.Round(cellWidthPx);

            // Create a simple placeholder image (1x1 transparent PNG)
            byte[] placeholderImage = CreatePlaceholderImage();

            // Add an OLE object with the placeholder image
            int oleIndex = worksheet.OleObjects.Add(rowIndex, columnIndex, height, width, placeholderImage);

            // Retrieve the OLE object and lock it
            OleObject oleObject = worksheet.OleObjects[oleIndex];
            oleObject.IsLocked = true;

            // Save the workbook
            string outputPath = "OleObjectForPowerPoint.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Returns a 1x1 transparent PNG byte array
    private static byte[] CreatePlaceholderImage()
    {
        // PNG file header for a 1x1 transparent image
        return new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x78,0x9C,0x63,0x00,0x01,0x00,0x00,
            0x05,0x00,0x01,0x0D,0x0A,0x2D,0xB4,0x00,
            0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
            0x42,0x60,0x82
        };
    }
}
