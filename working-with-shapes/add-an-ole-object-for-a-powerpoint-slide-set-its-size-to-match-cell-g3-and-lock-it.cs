// Title: Embed a PowerPoint slide as a locked OLE object sized to cell G3 using Aspose.Cells for .NET
// Common Searches: how to embed a .pptx slide as an OLE object in an Excel workbook with Aspose.Cells C# | set OLE object dimensions to match a specific worksheet cell (e.g., G3) in Aspose.Cells | lock an embedded OLE object to prevent moving or resizing in an Aspose.Cells spreadsheet | add a custom preview image for a PowerPoint OLE object in Excel using Aspose.Cells | retrieve pixel height and width of a cell for OLE object placement with Aspose.Cells
// Tags: embed PowerPoint slide as OLE object Aspose.Cells | size OLE object to cell dimensions Aspose.Cells | lock OLE object Aspose.Cells | display OLE object as icon Aspose.Cells | set FileFormatType.Pptx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AddOleObjectExample
{
    // The example creates a new workbook, calculates the pixel height and width of cell G3, reads a PNG preview image, adds a PowerPoint .pptx slide as an OLE object sized to those dimensions, displays it as an icon with a label, locks the object to prevent movement or resizing, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the resources. Adjust these paths to point to existing files on your machine.
            const string pptxPath = "SampleSlide.pptx";      // PowerPoint slide to embed
            const string previewImagePath = "SlidePreview.png"; // Small image that will be shown as the OLE icon
            const string outputPath = "WorkbookWithOleObject.xlsx";

            try
            {
                // Verify that the required files exist
                if (!File.Exists(pptxPath))
                    throw new FileNotFoundException($"PowerPoint file not found: {pptxPath}");
                if (!File.Exists(previewImagePath))
                    throw new FileNotFoundException($"Preview image file not found: {previewImagePath}");

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Determine the size of cell G3 (zero‑based indexes: row 2, column 6)
                // ------------------------------------------------------------
                int targetRow = 2;      // G3 -> row index 2
                int targetColumn = 6;   // G3 -> column index 6

                // Height of the row in pixels
                double rowHeightPixel = sheet.Cells.GetRowHeightPixel(targetRow);
                // Width of the column in pixels
                double columnWidthPixel = sheet.Cells.GetColumnWidthPixel(targetColumn);

                // Aspose.Cells expects integer pixel values for the OleObject size.
                int oleHeight = (int)Math.Round(rowHeightPixel);
                int oleWidth = (int)Math.Round(columnWidthPixel);

                // ------------------------------------------------------------
                // Read the preview image (icon) into a byte array.
                // This image is displayed in the worksheet when the OLE object is not opened.
                // ------------------------------------------------------------
                byte[] previewImageBytes = File.ReadAllBytes(previewImagePath);

                // ------------------------------------------------------------
                // Add the OLE object.
                // The Add method signature:
                //   int Add(int upperLeftRow, int upperLeftColumn, int height, int width, byte[] imageData)
                // ------------------------------------------------------------
                int oleIndex = sheet.OleObjects.Add(targetRow, targetColumn, oleHeight, oleWidth, previewImageBytes);
                OleObject oleObject = sheet.OleObjects[oleIndex];

                // ------------------------------------------------------------
                // Configure the OLE object to embed the PowerPoint slide.
                // ------------------------------------------------------------
                // Path to the PowerPoint file that will be embedded.
                oleObject.ObjectSourceFullName = Path.GetFullPath(pptxPath);

                // Set the file format type so that Excel knows how to handle the embedded data.
                oleObject.FileFormatType = FileFormatType.Pptx; // .pptx format

                // Display the object as an icon (the preview image we supplied).
                oleObject.DisplayAsIcon = true;

                // Optional: set a user‑friendly label that appears under the icon.
                oleObject.Label = "Embedded PowerPoint Slide";

                // ------------------------------------------------------------
                // Lock the OLE object so that users cannot move or resize it.
                // The Shape base class provides the IsLocked property.
                // ------------------------------------------------------------
                oleObject.IsLocked = true;

                // ------------------------------------------------------------
                // Save the workbook.
                // ------------------------------------------------------------
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while adding the OLE object:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
