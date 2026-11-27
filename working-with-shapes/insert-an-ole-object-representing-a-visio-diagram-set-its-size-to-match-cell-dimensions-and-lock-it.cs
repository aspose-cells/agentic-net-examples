// Title: Embedding a Visio diagram as a locked OLE object sized to a worksheet cell with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# insert Visio OLE object into Excel and lock it | set OLE object dimensions to match Excel cell size using Aspose.Cells | embed Visio drawing as OLE in .xlsx with optional preview image Aspose.Cells | prevent moving or resizing an OLE object in an Excel workbook via Aspose.Cells | specify ProgID Visio.Drawing for an OLE object when using Aspose.Cells
// Tags: insert visio ole object aspnet | size ole object to cell dimensions aspnet | lock ole object worksheet aspnet | set progid visio.drawing aspnet | add preview image to ole object aspnet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace InsertVisioOleObject
{
    // The example creates a new workbook, loads a Visio file and an optional PNG preview, adds an OLE object at cell B2, sizes it to the pixel width and height of that cell, assigns the Visio data to ObjectData, sets the ProgID to "Visio.Drawing", disables linking and icon display, locks the object to stop user resizing or moving, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the required files – adjust these paths to point to real files on your machine.
            const string visioFilePath = "diagram.vsdx";          // Visio diagram to embed
            const string previewImagePath = "preview.png";       // Optional preview image (can be empty)
            const string outputPath = "WorkbookWithVisioOle.xlsx";

            // Verify that the Visio file exists.
            if (!File.Exists(visioFilePath))
            {
                Console.WriteLine($"Visio file not found: {visioFilePath}");
                return;
            }

            // Load the preview image if it exists; otherwise use an empty byte array.
            byte[] previewImageData = File.Exists(previewImagePath)
                ? File.ReadAllBytes(previewImagePath)
                : new byte[0];

            // Load the Visio file data (the actual OLE object data).
            byte[] visioData = File.ReadAllBytes(visioFilePath);

            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Choose the cell that will be the upper‑left corner of the OLE object.
                // Example: cell B2 (row index 1, column index 1 – zero based).
                int startRow = 1;      // B2 row
                int startColumn = 1;   // B2 column

                // Determine the pixel width and height of a single cell.
                // Aspose.Cells provides helper methods for this.
                int cellWidthPixels = sheet.Cells.GetColumnWidthPixel(startColumn);
                int cellHeightPixels = sheet.Cells.GetRowHeightPixel(startRow);

                // Add the OLE object. The Add method signature is:
                // int Add(int upperLeftRow, int upperLeftColumn, int height, int width, byte[] imageData)
                int oleIndex = sheet.OleObjects.Add(startRow, startColumn, cellHeightPixels, cellWidthPixels, previewImageData);
                OleObject oleObject = sheet.OleObjects[oleIndex];

                // Set the embedded Visio data.
                oleObject.ObjectData = visioData;

                // Optional: set a ProgID for Visio. This helps Excel recognise the object type.
                // The ProgID for Visio drawings is usually "Visio.Drawing".
                oleObject.ProgID = "Visio.Drawing";

                // Ensure the object is embedded (not a link) and displayed directly, not as an icon.
                oleObject.IsLink = false;
                oleObject.DisplayAsIcon = false;

                // Lock the OLE object so that users cannot move or resize it in the UI.
                // The Shape base class (from which OleObject derives) exposes the IsLocked property.
                oleObject.IsLocked = true;

                // Save the workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The OLE object representing the Visio diagram has been inserted, sized to the cell, and locked.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while inserting the OLE object:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
