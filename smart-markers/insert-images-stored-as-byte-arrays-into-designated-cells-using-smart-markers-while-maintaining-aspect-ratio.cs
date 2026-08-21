// Title: Insert image byte array into a worksheet cell via smart markers with aspect‑ratio lock – Aspose.Cells for .NET
// Description: This C# sample reads a PNG file into a byte[], adds it to a DataTable, places the &IMG smart marker in cell A1, processes the marker with WorkbookDesigner, then configures the inserted picture to stay inside the cell and retain its proportions before saving the workbook.
// Keywords: Aspose.Cells image byte array | smart marker &IMG | place picture in cell | lock aspect ratio | IsPlacedInCell property | C# Excel image insertion | WorkbookDesigner image handling
// Common Searches: Aspose.Cells insert image from byte[] using smart marker | C# keep picture aspect ratio when adding to Excel with Aspose | How to embed PNG into a specific cell with Aspose.Cells | Set IsPlacedInCell true for smart‑marker images | Smart marker image scaling Aspose.Cells
// Developer Intent: Add a picture stored as a byte[] to a designated Excel cell through a smart marker, ensuring the image remains within the cell boundaries and preserves its original proportions.
// Use Cases: Load a company logo into a byte array, bind it to a DataTable, use the &IMG marker in a template cell, and lock the aspect ratio so the logo fits neatly. | Generate product catalogs where each row contains a product photo; populate a DataTable with multiple byte[] columns and apply the same smart‑marker logic to insert and size each image. | Create a reporting template that automatically places user‑uploaded photos into predefined cells, preventing distortion and keeping the images confined to their cells.
// AI Prompts: Provide C# code that inserts multiple images from a DataTable into consecutive rows using smart markers, and sets each picture to be placed in the cell with a locked aspect ratio. | Explain how to automatically adjust row height and column width after inserting an image with IsPlacedInCell enabled in Aspose.Cells. | Show how to handle different image formats (PNG, JPEG, GIF) stored as byte arrays when using smart markers to embed them in a worksheet.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsImageSmartMarkerDemo
{
    // This C# sample reads a PNG file into a byte[], adds it to a DataTable, places the &IMG smart marker in cell A1, processes the marker with WorkbookDesigner, then configures the inserted picture to stay inside the cell and retain its proportions before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- 1. Prepare image byte array ----------
                const string imagePath = "sample.png";
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file \"{imagePath}\" not found.");
                    return;
                }

                byte[] imageBytes = File.ReadAllBytes(imagePath);

                // ---------- 2. Create a data source containing the image ----------
                DataTable dt = new DataTable("Images");
                dt.Columns.Add("ImageData", typeof(byte[]));
                dt.Rows.Add(imageBytes);

                // ---------- 3. Create a workbook and place a smart marker ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                // Put the smart marker in cell A1.
                sheet.Cells["A1"].PutValue("&IMG");

                // ---------- 4. Process the smart marker ----------
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);
                designer.Process();

                // ---------- 5. Adjust the inserted picture ----------
                if (sheet.Pictures.Count > 0)
                {
                    Picture pic = sheet.Pictures[0];
                    // Place the picture inside the cell (maintains cell boundaries).
                    pic.IsPlacedInCell = true;
                    // Lock aspect ratio so the image scales proportionally.
                    pic.IsAspectRatioLocked = true;
                }

                // ---------- 6. Save the workbook ----------
                const string outputPath = "ImageSmartMarkerOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
