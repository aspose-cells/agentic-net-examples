// Title: Embed a PowerPoint slide as a locked OLE object sized to cell G3 using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells embed PowerPoint slide as OLE object and lock it in Excel | how to set OLE object dimensions to match a specific worksheet cell with Aspose.Cells | Aspose.Cells add locked OLE object at G3 cell using pixel sizing | retrieve column width and row height in pixels for a given cell in Aspose.Cells C# | save workbook after embedding a PPTX as an OLE object with Aspose.Cells .NET
// Tags: embed PowerPoint OLE object Aspose.Cells | size OLE object using cell pixel dimensions | lock OLE object in Excel workbook | OleObject.Add method pixel sizing | get column width pixel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for OleObject

namespace Example
{
    // The example creates a new workbook, accesses the first worksheet, checks for the PowerPoint file, reads it into a byte array, obtains the pixel width of column G and the pixel height of row 3, adds the PowerPoint slide as an OLE object positioned at cell G3 with matching dimensions, locks the OLE object to prevent moving or resizing, and saves the workbook as output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Path to the PowerPoint file that will be embedded as an OLE object
                string pptPath = "slide.pptx";

                // Verify the PowerPoint file exists
                if (!File.Exists(pptPath))
                {
                    Console.WriteLine($"File not found: {pptPath}");
                    return;
                }

                // Zero‑based indexes for cell G3 (column G = 6, row 3 = 2)
                int rowIndex = 2;
                int columnIndex = 6;

                // Retrieve the pixel width of column G and the pixel height of row 3
                int cellWidth = sheet.Cells.GetColumnWidthPixel(columnIndex);
                int cellHeight = sheet.Cells.GetRowHeightPixel(rowIndex);

                // Read the PowerPoint file into a byte array
                byte[] oleData = File.ReadAllBytes(pptPath);

                // Add the OLE object (PowerPoint slide) at cell G3 with the same size as the cell
                int oleIndex = sheet.OleObjects.Add(rowIndex, columnIndex, cellHeight, cellWidth, oleData);
                OleObject ole = sheet.OleObjects[oleIndex];

                // Lock the OLE object to prevent moving or resizing
                ole.IsLocked = true;

                // Save the workbook to a file
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
