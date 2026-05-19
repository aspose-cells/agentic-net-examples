using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ExportOleObjectsToPdfApp
{
    class ExportOleObjectsToPdf
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Paths for the icon image and the embedded OLE file
                string iconPath = "icon.png";
                string olePath = "sample.xlsx";

                // Verify that the required files exist
                if (!File.Exists(iconPath))
                {
                    Console.WriteLine($"Icon file not found: {iconPath}");
                    return;
                }
                if (!File.Exists(olePath))
                {
                    Console.WriteLine($"OLE source file not found: {olePath}");
                    return;
                }

                // Load file bytes
                byte[] iconData = File.ReadAllBytes(iconPath);
                byte[] oleData = File.ReadAllBytes(olePath);

                // Add an OLE object to the worksheet using the icon image
                int oleIndex = sheet.OleObjects.Add(5, 2, 200, 200, iconData);
                OleObject ole = sheet.OleObjects[oleIndex];
                ole.ObjectData = oleData;
                ole.DisplayAsIcon = true;          // Show as an icon
                ole.Label = "Sample Excel File";   // Icon label

                // Show placeholders instead of the actual OLE objects
                workbook.Settings.DisplayDrawingObjects = DisplayDrawingObjects.Placeholders;

                // Configure PDF save options (do not embed attachments)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = false
                };

                // Save the workbook as PDF
                string outputPdf = "OleObjectsPlaceholders.pdf";
                workbook.Save(outputPdf, pdfOptions);
                Console.WriteLine($"PDF saved to {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}