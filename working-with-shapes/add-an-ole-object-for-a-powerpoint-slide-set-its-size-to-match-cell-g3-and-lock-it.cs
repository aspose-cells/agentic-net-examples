// Title: Insert a PowerPoint slide as a locked OLE object sized to cell G3 with Aspose.Cells for .NET
// AI Prompts: Read a .pptx file into a byte array and embed it as an OLE object at worksheet cell G3, then set its width and height to match the cell and lock the object using Aspose.Cells C#. | Create an OLE object from PowerPoint data, position it on row 3 column G, adjust its dimensions to the cell's pixel size, and enable the IsLocked property in Aspose.Cells. | Programmatically add a PowerPoint slide to an Excel worksheet, resize the OLE object to the exact size of cell G3, and prevent editing by locking the object with the Aspose.Cells API.
// Common Searches: aspocells embed powerpoint slide as oleobject in specific cell | c# set oleobject dimensions to match excel cell size | how to lock an oleobject in an Aspose.Cells worksheet | add oleobject from byte array using Aspose.Cells .NET | resize oleobject to cell G3 pixel dimensions Aspose.Cells
// Tags: oleobject insertion from binary data | resize oleobject to match cell dimensions | prevent editing of oleobject in worksheet | powerpoint oleobject embedding | aspocells oleobject API usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for OleObject

namespace AsposeCellsExample
{
    // The example loads an Excel workbook and a PowerPoint file, embeds the PowerPoint slide as an OLE object positioned in cell G3, resizes it to the cell's pixel dimensions, locks the object to prevent editing, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string pptPath = "slide.pptx";
                string outputPath = "output.xlsx";

                // Verify required files exist
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input workbook not found: {inputPath}");
                    return;
                }
                if (!File.Exists(pptPath))
                {
                    Console.WriteLine($"PowerPoint file not found: {pptPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Target cell G3 (zero‑based indices)
                int colIndex = 6; // G
                int rowIndex = 2; // 3

                // Get cell dimensions in pixels (cast to int)
                int cellWidth = (int)sheet.Cells.GetColumnWidthPixel(colIndex);
                int cellHeight = (int)sheet.Cells.GetRowHeightPixel(rowIndex);

                // Read PowerPoint file into a byte array (required by OleObjects.Add overload)
                byte[] pptData = File.ReadAllBytes(pptPath);

                // Add OLE object (PowerPoint slide) at G3 with the cell size
                int oleIndex = sheet.OleObjects.Add(rowIndex, colIndex, cellHeight, cellWidth, pptData, "PowerPoint.Show");
                OleObject ole = sheet.OleObjects[oleIndex];
                ole.IsLocked = true; // Prevent editing

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
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
