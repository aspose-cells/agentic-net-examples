// Title: Insert Images into Excel with Aspose.Cells Smart Markers Using a byte[] Source (C#)
// Description: C# sample that demonstrates how to embed a PNG image into an Excel workbook via Aspose.Cells Smart Markers. A smart‑marker "&=Image:Photo" is placed in cell A1, a DataTable with a byte[] column supplies the image bytes, WorkbookDesigner processes the marker, and the file is saved as an XLSX workbook.
// Keywords: Aspose.Cells | Smart Markers | C# image insertion | byte[] image Excel | WorkbookDesigner | Excel picture from database | populate Excel with pictures | Aspose.Cells &=Image tag | Excel automation C# | smart marker image example
// Common Searches: Aspose.Cells insert image from byte array | C# smart marker image tag example | how to use &=Image smart marker in Aspose.Cells | load PNG into byte[] and add to Excel with Aspose | Aspose.Cells WorkbookDesigner image column
// Developer Intent: Add an image to a worksheet by linking a byte[] data column to a smart‑marker image tag.
// Use Cases: Generate a product catalog where each product row displays a photo stored as a BLOB. | Create an employee directory workbook that embeds staff portraits from a byte[] field. | Automate a sales report that inserts chart screenshots saved as byte arrays into specific cells.
// AI Prompts: Show how to modify the code to insert multiple images from a DataTable with several rows. | Explain how to set the width and height of images inserted via the &=Image smart marker. | Provide an example that reads an image into a MemoryStream before assigning it to the byte[] column.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerImageDemo
{
    // C# sample that demonstrates how to embed a PNG image into an Excel workbook via Aspose.Cells Smart Markers. A smart‑marker "&=Image:Photo" is placed in cell A1, a DataTable with a byte[] column supplies the image bytes, WorkbookDesigner processes the marker, and the file is saved as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Verify that the image file exists to avoid FileNotFoundException
                const string imagePath = "sample.png";
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file \"{imagePath}\" not found.");
                    return;
                }

                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Insert a smart marker that expects an image.
                // The syntax "&=Image:Photo" tells the processor to replace this cell with an image
                // taken from the "Photo" column of the data source.
                sheet.Cells["A1"].PutValue("&=Image:Photo");

                // Prepare a DataTable as the data source.
                // The "Photo" column must be of type byte[] and contain the image data.
                DataTable dt = new DataTable("Products");
                dt.Columns.Add("Photo", typeof(byte[]));

                // Load the image file into a byte array.
                byte[] imageBytes = File.ReadAllBytes(imagePath);
                dt.Rows.Add(imageBytes);

                // Process the workbook with the data source using WorkbookDesigner (Smart Marker processor)
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);
                designer.Process();

                // Save the workbook (lifecycle: save)
                const string outputPath = "SmartMarkerImageOutput.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
