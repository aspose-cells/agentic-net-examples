// Title: Insert Images from Byte Arrays into Cells Using Smart Markers and Keep Aspect Ratio – Aspose.Cells for .NET
// Description: Demonstrates how to read a PNG file into a byte[] array, bind it to a DataTable, place a smart marker "&=Image" in a target cell, process the marker with WorkbookDesigner, lock the picture's aspect ratio, attach the image to the cell, and save the workbook.
// Keywords: Aspose.Cells image from byte array | smart markers insert picture .NET | lock picture aspect ratio Aspose.Cells | WorkbookDesigner image insertion | C# embed image in Excel cell | place picture in cell with smart marker
// Common Searches: how to embed a byte[] image in Excel using Aspose.Cells smart markers | preserve aspect ratio when inserting images via Aspose.Cells | set picture to move with cell after smart marker processing | insert multiple images from DataTable using Aspose.Cells smart markers | C# Aspose.Cells image insertion example
// Developer Intent: Embed a byte‑array image into a specific worksheet cell through a smart marker and ensure the picture retains its original proportions.
// Use Cases: Generate product catalogs where each product row displays a photo stored as a byte[] in a database. | Create employee directories that pull profile pictures from a byte[] field and embed them in designated cells. | Add a company logo from a byte[] to an invoice header while keeping the logo's dimensions proportional.
// AI Prompts: Show code to insert several images from a DataTable into consecutive rows using smart markers, with each picture locked to its aspect ratio. | Provide an example of using WorkbookDesigner with a List<T> where T contains a byte[] Image property for smart‑marker population. | Explain how to adjust a picture's size relative to its cell after locking the aspect ratio in Aspose.Cells.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsImageSmartMarkerDemo
{
    // Demonstrates how to read a PNG file into a byte[] array, bind it to a DataTable, place a smart marker "&=Image" in a target cell, process the marker with WorkbookDesigner, lock the picture's aspect ratio, attach the image to the cell, and save the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define the target cell where the image will be placed using a smart marker
                // The smart marker syntax "&=Image" tells Aspose.Cells to treat the cell value as an image
                Cell markerCell = sheet.Cells["B2"];
                markerCell.PutValue("&=Image");

                // Prepare image data as a byte array (example uses a local PNG file)
                byte[] imageBytes = null;
                string imagePath = "sample.png";
                if (File.Exists(imagePath))
                {
                    imageBytes = File.ReadAllBytes(imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping image insertion.");
                }

                // Create a DataTable that will serve as the data source for the smart marker
                DataTable dt = new DataTable();
                dt.Columns.Add("Image", typeof(byte[]));
                dt.Rows.Add(imageBytes);

                // Process the smart markers using WorkbookDesigner (compatible with all Aspose.Cells versions)
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);
                designer.Process();

                // After processing, the image is placed in the cell as an embedded picture.
                // Retrieve the picture object that was created and lock its aspect ratio.
                // The picture is the last one added to the worksheet's Pictures collection.
                if (sheet.Pictures.Count > 0)
                {
                    Picture pic = sheet.Pictures[sheet.Pictures.Count - 1];
                    pic.IsAspectRatioLocked = true; // Maintain original aspect ratio
                    pic.IsPlacedInCell = true;      // Ensure the picture moves/resizes with the cell
                }

                // Save the workbook
                string outputPath = "ImageSmartMarkerOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
