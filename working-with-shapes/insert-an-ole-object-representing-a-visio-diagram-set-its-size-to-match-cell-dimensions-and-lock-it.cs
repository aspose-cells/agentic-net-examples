// Title: Insert a Visio .vsdx diagram as a locked OLE object sized to a specific Excel cell using Aspose.Cells for .NET (C#)
// AI Prompts: Add a Visio .vsdx file as an OLE object to a worksheet cell, use the cell's pixel width and height for sizing, and lock the object with Aspose.Cells for .NET. | Change the target row, column, or Visio source file while automatically resizing and locking the OLE object to the cell dimensions using Aspose.Cells.
// Common Searches: C# Aspose.Cells embed Visio .vsdx as locked OLE object in a specific Excel cell | How to size an OLE object to match Excel cell pixel dimensions with Aspose.Cells | Lock OLE object after inserting into worksheet using Aspose.Cells for .NET | Get column width in pixels Aspose.Cells C# for OLE object sizing
// Tags: add visio oleobject to worksheet cell Aspose.Cells | set oleobject dimensions using GetColumnWidthPixel Aspose.Cells | lock oleobject in Excel workbook Aspose.Cells | embed vsdx file as oleobject C# Aspose.Cells | match oleobject size to cell pixel size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // For OleObject

namespace AsposeCellsExample
{
    // The example creates a new workbook, calculates the pixel width and height of a target cell, reads a Visio .vsdx file into a byte array, inserts it as an OLE object at the specified row and column using those dimensions, locks the OLE object to prevent moving or resizing, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Target cell (zero‑based indices) where the OLE object will be placed
                int targetRow = 5;
                int targetColumn = 2;

                // Get pixel dimensions of the target cell
                int cellWidth = (int)Math.Ceiling((double)sheet.Cells.GetColumnWidthPixel(targetColumn));
                int cellHeight = (int)Math.Ceiling((double)sheet.Cells.GetRowHeightPixel(targetRow));

                // Load Visio file bytes if the file exists; otherwise use an empty placeholder
                byte[] oleData;
                string visioFilePath = "sample.vsdx";

                if (File.Exists(visioFilePath))
                {
                    try
                    {
                        oleData = File.ReadAllBytes(visioFilePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to read Visio file: {ex.Message}");
                        oleData = Array.Empty<byte>();
                    }
                }
                else
                {
                    oleData = Array.Empty<byte>();
                }

                // Insert the OLE object. The Add method returns the index of the newly added object.
                OleObject visioOle = null;
                try
                {
                    int oleIndex = sheet.OleObjects.Add(targetRow, targetColumn, cellHeight, cellWidth, oleData);
                    visioOle = sheet.OleObjects[oleIndex];
                    // Lock the OLE object to prevent moving or resizing
                    visioOle.IsLocked = true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add OLE object: {ex.Message}");
                }

                // Save the workbook
                string outputPath = "VisioDiagram.xlsx";
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
