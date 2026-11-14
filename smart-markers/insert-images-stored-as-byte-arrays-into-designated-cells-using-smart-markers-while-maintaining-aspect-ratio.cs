// Title: Insert images from byte[] into Excel cells with Aspose.Cells smart markers while preserving aspect ratio (C#)
// Common Searches: C# Aspose.Cells insert picture from byte array using smart markers | how to keep image aspect ratio when adding pictures via Aspose.Cells smart markers | using WorkbookDesigner to bind a DataTable of byte[] images to Excel smart markers | set row height and column width automatically for images inserted with Aspose.Cells smart markers | Aspose.Cells smart marker syntax for embedding images stored as byte arrays
// Tags: byte array image insertion smart markers | Aspose.Cells maintain picture aspect ratio | WorkbookDesigner bind DataTable images | auto adjust row height for embedded pictures | C# Excel smart marker image embedding

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, places a smart marker "&=\"Image\"" in cell A1, loads PNG and JPG files into a DataTable as byte arrays, sets the table as the data source for a WorkbookDesigner, processes the smart markers to embed the images, adjusts row heights and column width for proper display, and saves the result as output.xlsx while preserving the original aspect ratio of each picture.
class InsertImagesWithSmartMarkers
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Place a smart marker for an image in cell A1
            sheet.Cells["A1"].PutValue("&=\"Image\"");

            // Prepare a DataTable that will serve as the data source for the smart marker
            DataTable dt = new DataTable();
            dt.Columns.Add("Image", typeof(byte[]));

            // Load image bytes if files exist
            string[] imageFiles = { "image1.png", "image2.jpg" };
            foreach (string imgPath in imageFiles)
            {
                if (File.Exists(imgPath))
                {
                    byte[] imgBytes = File.ReadAllBytes(imgPath);
                    dt.Rows.Add(imgBytes);
                }
                else
                {
                    Console.WriteLine($"Warning: Image file '{imgPath}' not found. Skipping.");
                }
            }

            // Ensure at least one row exists to avoid processing errors
            if (dt.Rows.Count == 0)
            {
                Console.WriteLine("No images loaded. Exiting.");
                return;
            }

            // Create a WorkbookDesigner to process smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("DataSource", dt);
            designer.Process();

            // Adjust row heights and column widths for better display
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                sheet.Cells.SetRowHeight(i, 100);
            }
            sheet.Cells.SetColumnWidth(0, 30);

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
