// Title: Insert a PNG byte array into an Excel cell with Aspose.Cells smart markers while preserving aspect ratio
// AI Prompts: Generate C# code that binds a DataTable containing a byte[] image column to a WorkbookDesigner smart marker and inserts the picture into a specific worksheet cell. | Show how to lock the aspect ratio of a picture placed by a smart marker and set its width and height scaling using Aspose.Cells APIs. | Demonstrate converting a Base64‑encoded PNG to a byte array, adding it to a DataTable, and processing the smart marker to embed the image in an .xlsx file.
// Common Searches: asp.net insert image from byte array into excel using smart markers | keep picture aspect ratio when using Aspose.Cells WorkbookDesigner | bind image byte[] column to smart marker in C# Aspose.Cells example | scale picture inserted by smart marker Aspose.Cells C# | convert base64 png to byte array for Aspose.Cells smart marker
// Tags: smart marker image insertion byte array Aspose.Cells | lock picture aspect ratio WorkbookDesigner | picture width height scaling Aspose.Cells | base64 png to byte[] for Excel smart marker | C# Aspose.Cells embed image from DataTable

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, places a smart marker '&=Image' in cell A1, converts a Base64‑encoded PNG to a byte array, adds it to a DataTable, sets the table as the data source for WorkbookDesigner, processes the smart marker to embed the image, then locks the picture's aspect ratio and applies 100% width and height scaling before saving the file as SmartMarkerImage.xlsx.
class InsertImageSmartMarker
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a smart marker that expects an image into cell A1
            sheet.Cells["A1"].PutValue("&=Image");

            // Sample image as a Base64 encoded 1x1 pixel PNG
            const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO2b4ZcAAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(pngBase64);

            // Prepare a DataTable with a byte[] column for the image
            DataTable dt = new DataTable();
            dt.Columns.Add("Image", typeof(byte[]));
            DataRow dr = dt.NewRow();
            dr["Image"] = imageBytes;
            dt.Rows.Add(dr);

            // Process the smart marker using WorkbookDesigner (correct API)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process();

            // Adjust properties of the picture inserted by the smart marker
            foreach (Picture pic in sheet.Pictures)
            {
                if (pic.IsPlacedInCell) // Picture inserted via smart marker will be placed in a cell
                {
                    pic.IsAspectRatioLocked = true; // Keep original aspect ratio
                    pic.WidthScale = 100;            // 100% width scaling
                    pic.HeightScale = 100;           // 100% height scaling
                }
            }

            // Save the workbook with the embedded image
            string outputPath = "SmartMarkerImage.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
