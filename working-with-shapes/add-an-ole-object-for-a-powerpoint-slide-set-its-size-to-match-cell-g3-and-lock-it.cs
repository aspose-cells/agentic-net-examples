using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Cell G3 corresponds to column index 6 (zero‑based) and row index 2 (zero‑based)
            int colIndex = 6;
            int rowIndex = 2;

            // Get cell dimensions in pixels using the built‑in helpers
            int widthPixels = sheet.Cells.GetColumnWidthPixel(colIndex);
            int heightPixels = sheet.Cells.GetRowHeightPixel(rowIndex);

            // Add an OLE object at G3 with the same size as the cell.
            // Use an empty byte array for the image data; replace with actual image bytes if needed.
            byte[] imageData = new byte[0];
            int oleIndex = sheet.OleObjects.Add(rowIndex, colIndex, heightPixels, widthPixels, imageData);

            // Retrieve the added OLE object and lock it so it cannot be modified when the sheet is protected.
            OleObject ole = sheet.OleObjects[oleIndex];
            ole.IsLocked = true;

            // Save the workbook
            string outputPath = "OleObjectInSlide.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}