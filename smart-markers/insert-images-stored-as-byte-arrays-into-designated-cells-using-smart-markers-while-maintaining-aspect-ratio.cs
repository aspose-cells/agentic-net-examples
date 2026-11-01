// Title: Insert a PNG image stored as a byte array into a specific Excel cell using Aspose.Cells smart markers while preserving aspect ratio (C#)
// Common Searches: C# Aspose.Cells insert image from byte[] into worksheet cell using smart markers | How to bind a DataTable that contains image byte arrays to a smart marker in Aspose.Cells | Preserve picture aspect ratio when inserting images with Aspose.Cells WorkbookDesigner | Insert PNG byte array into Excel using the &amp;=Image smart marker | Adjust picture size after processing smart markers in Aspose.Cells C#
// Tags: smart marker image insertion byte array | Aspose.Cells preserve aspect ratio | WorkbookDesigner bind datatable image column | C# insert png into excel cell | adjust picture dimensions after smart marker

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsImageSmartMarker
{
    // Class that holds the image data as a byte array.
    // The example reads a PNG file into a byte array, places the array into a DataTable, adds the smart marker "&=Image" to cell A1, binds the DataTable to a WorkbookDesigner, processes the smart marker to embed the image, optionally sets the picture height (letting the width adjust automatically to maintain the original aspect ratio), and saves the workbook as OutputWithImage.xlsx.
    public class ImageData
    {
        public byte[] Image { get; set; } = Array.Empty<byte>();
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Prepare image byte array (ensure the file exists)
                // -----------------------------------------------------------------
                string imagePath = @"C:\Images\sample.png";
                byte[] imageBytes = Array.Empty<byte>();

                if (File.Exists(imagePath))
                {
                    imageBytes = File.ReadAllBytes(imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // -----------------------------------------------------------------
                // 2. Create a DataTable that contains the image byte array
                // -----------------------------------------------------------------
                DataTable dt = new DataTable();
                dt.Columns.Add("Image", typeof(byte[]));
                dt.Rows.Add(imageBytes);

                // -----------------------------------------------------------------
                // 3. Create a new workbook and place a smart marker for the image
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();                     // create a new workbook
                Worksheet sheet = workbook.Worksheets[0];              // get the first worksheet
                sheet.Cells["A1"].PutValue("&=Image");                  // smart marker for image insertion

                // -----------------------------------------------------------------
                // 4. Use WorkbookDesigner to process the smart marker with the data source
                // -----------------------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);                             // bind the DataTable as data source
                designer.Process();                                    // process smart markers

                // -----------------------------------------------------------------
                // 5. Adjust the inserted picture (optional)
                // -----------------------------------------------------------------
                if (sheet.Pictures.Count > 0)
                {
                    Picture picture = sheet.Pictures[0];
                    // The Picture class in this version does not expose an IsLockedAspectRatio property.
                    // If needed, you can manually set Height or Width while preserving the aspect ratio.
                    // Example: picture.Height = 100; // Width will adjust automatically based on original ratio.
                }

                // -----------------------------------------------------------------
                // 6. Save the workbook
                // -----------------------------------------------------------------
                string outputPath = "OutputWithImage.xlsx";

                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
