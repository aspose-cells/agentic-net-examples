// Title: Embed a Visio diagram as a locked OLE object sized to a cell range using Aspose.Cells for .NET
// Description: Creates a new workbook, loads a PNG preview and a .vsdx Visio file, adds an OLE object anchored at cell B2, embeds the Visio diagram, resizes it to cover cells B2:C5 based on default row height and column width, locks the object to prevent movement, and saves the file as VisioOleObject.xlsx.
// Keywords: Aspose.Cells | C# OLE object | Visio embed Excel | lock OLE object | resize OLE to cell range | Excel automation | Visio diagram in workbook | programmatic OLE sizing
// Common Searches: Aspose.Cells embed Visio as OLE object | Resize OLE object to specific cells in .NET | Lock OLE objects in Excel worksheets using Aspose | Calculate OLE width height from column width row height | Add Visio diagram to Excel with Aspose.Cells C#
// Developer Intent: Insert a Visio diagram as an embedded OLE object, adjust its dimensions to match a target cell range, and lock it to keep the layout stable.
// Use Cases: Automated report generation where each sheet contains a fixed‑position Visio flowchart aligned with a data table. | Template creation that places a Visio diagram into a predefined cell block and locks it to preserve layout across users. | Batch processing of workbooks to embed different Visio files into designated ranges while preventing accidental edits.
// AI Prompts: Generate code that computes OLE object Width and Height from actual column widths and row heights instead of pixel approximations. | Show how to add multiple Visio OLE objects to different cell ranges and lock each one programmatically. | Explain how to maintain the aspect ratio of the Visio preview image while fitting the OLE object within a specified cell block.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Creates a new workbook, loads a PNG preview and a .vsdx Visio file, adds an OLE object anchored at cell B2, embeds the Visio diagram, resizes it to cover cells B2:C5 based on default row height and column width, locks the object to prevent movement, and saves the file as VisioOleObject.xlsx.
class InsertVisioOleObject
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell where the OLE object will be anchored (e.g., B2)
            int topRow = 1;        // Row index (zero‑based) -> B2
            int leftColumn = 1;    // Column index (zero‑based) -> B2

            // Load preview image and Visio file if they exist
            string previewPath = "visio_preview.png";
            string visioPath = "diagram.vsdx";

            byte[] previewImage = null;
            byte[] visioData = null;

            if (File.Exists(previewPath))
            {
                previewImage = File.ReadAllBytes(previewPath);
            }
            else
            {
                Console.WriteLine($"Preview image not found: {previewPath}");
            }

            if (File.Exists(visioPath))
            {
                visioData = File.ReadAllBytes(visioPath);
            }
            else
            {
                Console.WriteLine($"Visio file not found: {visioPath}");
            }

            // Proceed only if both files were loaded successfully
            if (previewImage != null && visioData != null)
            {
                // Add the OLE object with the preview image.
                // Height and width are set to 0 for now; they will be adjusted later.
                int oleIndex = worksheet.OleObjects.Add(topRow, leftColumn, 0, 0, previewImage);
                OleObject oleObject = worksheet.OleObjects[oleIndex];

                // Embed the Visio file into the OLE object.
                // linkToFile = false (embed), displayAsIcon = false, label = "Visio Diagram"
                oleObject.SetEmbeddedObject(false, visioData, "diagram.vsdx", false, "Visio Diagram");

                // ------------------------------------------------------------
                // Resize the OLE object to match a specific cell range.
                // Example: make it cover cells B2:C5 (rows 1‑4, columns 1‑2)
                // ------------------------------------------------------------
                int endRow = 4;      // Row index for row 5
                int endColumn = 2;   // Column index for column C

                // Approximate pixel size: default row height ≈ 20 px, default column width ≈ 64 px
                int rowsCovered = endRow - topRow + 1;
                int colsCovered = endColumn - leftColumn + 1;
                int pixelHeight = rowsCovered * 20;   // height in pixels
                int pixelWidth = colsCovered * 64;    // width in pixels

                oleObject.Height = pixelHeight;
                oleObject.Width = pixelWidth;

                // Lock the OLE object so it cannot be moved or resized when the sheet is protected
                oleObject.IsLocked = true;
            }

            // Save the workbook
            workbook.Save("VisioOleObject.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
