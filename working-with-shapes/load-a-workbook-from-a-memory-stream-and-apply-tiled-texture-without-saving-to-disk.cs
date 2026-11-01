// Title: Load an Excel workbook from a MemoryStream and apply a solid fill style to simulate a tiled background using Aspose.Cells for .NET
// Common Searches: asp.net load excel workbook from byte array using Aspose.Cells | apply background style to all cells in Aspose.Cells without saving the file | create in‑memory worksheet style with solid fill Aspose.Cells C# | how to use MemoryStream with Aspose.Cells to modify a workbook | simulate tiled texture in Excel using Aspose.Cells style flags
// Tags: load workbook from memorystream Aspose.Cells | apply solid fill style to worksheet Aspose.Cells | in‑memory Excel manipulation C# | style all cells background Aspose.Cells | tiled texture simulation Excel Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example reads Excel workbook bytes, loads them into a MemoryStream, creates a Workbook object, accesses the first worksheet, defines a solid fill style (white foreground, light gray background) to mimic a tiled texture, applies this style to every cell, and completes the operation entirely in memory without writing any file to disk.
class Program
{
    static void Main()
    {
        try
        {
            // Obtain the workbook bytes from a source (e.g., database, network, etc.).
            byte[] workbookBytes = GetWorkbookBytes();

            if (workbookBytes == null || workbookBytes.Length == 0)
                throw new InvalidOperationException("Workbook byte array is empty.");

            // Load the workbook from a memory stream (no file system access required).
            using (MemoryStream memoryStream = new MemoryStream(workbookBytes))
            {
                Workbook workbook = new Workbook(memoryStream);

                // Access the first worksheet (adjust index as needed).
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a style with a solid fill (texture fill not supported in this version).
                Style tiledStyle = workbook.CreateStyle();
                tiledStyle.Pattern = BackgroundType.Solid;          // Solid fill pattern.
                tiledStyle.ForegroundColor = Color.White;          // Main cell background.
                tiledStyle.BackgroundColor = Color.LightGray;      // Secondary color (acts like a texture).

                // Apply the style to all cells in the worksheet.
                StyleFlag styleFlag = new StyleFlag { All = true };
                worksheet.Cells.ApplyStyle(tiledStyle, styleFlag);
            }

            // No saving to disk is performed as per the requirement.
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Placeholder method to obtain workbook bytes.
    // Replace with actual implementation (e.g., reading from a database, network, etc.).
    static byte[] GetWorkbookBytes()
    {
        // Example: create an empty workbook and return its bytes.
        using (MemoryStream ms = new MemoryStream())
        {
            Workbook tempWb = new Workbook();
            tempWb.Save(ms, SaveFormat.Xlsx);
            return ms.ToArray();
        }
    }
}
