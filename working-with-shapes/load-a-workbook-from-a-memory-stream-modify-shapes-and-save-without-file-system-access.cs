// Title: C# – Load, Edit, and Save Excel Shapes via MemoryStream with Aspose.Cells
// Description: Demonstrates how to create an Excel workbook in memory, add a rectangle shape, reload the workbook from a MemoryStream, change the shape’s name and fill color, and save the result back to another MemoryStream—entirely without touching the file system.
// Keywords: Aspose.Cells C# MemoryStream | edit Excel shape in memory | modify shape name Aspose.Cells | change shape fill color .NET | load workbook from stream | save workbook to stream | in‑memory Excel processing | no disk I/O Aspose.Cells | shape manipulation Aspose.Cells
// Common Searches: Aspose.Cells edit shape without saving to disk | C# load Excel from MemoryStream and change rectangle color | save modified workbook to MemoryStream Aspose.Cells | how to rename a shape in an Excel file using Aspose.Cells | in‑memory Excel shape update C#
// Developer Intent: Load an Excel workbook from a MemoryStream, update a shape’s name and fill color, and write the modified workbook to another MemoryStream, avoiding any file‑system operations.
// Use Cases: Web API that receives an uploaded XLSX, updates chart or diagram shapes on the fly, and returns the altered file as a byte array. | Automated report generator that builds Excel files entirely in memory, customizes shapes, and emails the result without creating temporary files. | Batch processing of workbooks stored in a database where shape properties are refreshed using streams to eliminate disk I/O.
// AI Prompts: Provide C# code that loads an Excel workbook from a byte array, changes the fill color of every rectangle shape to red, and returns the updated workbook as a byte array using Aspose.Cells. | Show an example that reads a workbook from a MemoryStream, prefixes "Updated_" to each shape’s name, and writes the workbook to a new MemoryStream. | Explain best practices for efficiently modifying shape properties in large numbers of in‑memory workbooks with Aspose.Cells while avoiding any file‑system access.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create an Excel workbook in memory, add a rectangle shape, reload the workbook from a MemoryStream, change the shape’s name and fill color, and save the result back to another MemoryStream—entirely without touching the file system.
class ShapeMemoryDemo
{
    static void Main()
    {
        // Create an initial workbook in memory and add a rectangle shape
        using (MemoryStream sourceStream = new MemoryStream())
        {
            Workbook wb = new Workbook();                                   // create workbook
            Worksheet ws = wb.Worksheets[0];
            Shape rect = ws.Shapes.AddRectangle(1, 1, 0, 0, 100, 100);      // add rectangle
            rect.Name = "OriginalRect";
            rect.FillFormat.ForeColor = System.Drawing.Color.Blue;          // set initial color

            // Save the workbook to the memory stream (XLSX format)
            wb.Save(sourceStream, SaveFormat.Xlsx);
            sourceStream.Position = 0;                                      // reset for reading

            // Load the workbook from the memory stream
            Workbook loadedWb = new Workbook(sourceStream);
            Worksheet loadedWs = loadedWb.Worksheets[0];

            // Modify the first shape: change its name and fill color
            if (loadedWs.Shapes.Count > 0)
            {
                Shape shape = loadedWs.Shapes[0];
                shape.Name = "ModifiedRect";
                shape.FillFormat.ForeColor = System.Drawing.Color.Green;
            }

            // Save the modified workbook to another memory stream
            using (MemoryStream resultStream = new MemoryStream())
            {
                loadedWb.Save(resultStream, SaveFormat.Xlsx);
                // resultStream now holds the updated workbook; its length can be inspected
                Console.WriteLine($"Modified workbook size: {resultStream.Length} bytes");
            }
        }
    }
}
