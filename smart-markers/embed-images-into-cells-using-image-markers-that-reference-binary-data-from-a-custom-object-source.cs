// Title: How to embed a PNG image into an Excel cell using Aspose.Cells smart image markers and a byte[] data source in C#
// AI Prompts: Write C# code that reads a PNG file into a byte array, adds it to a DataTable, and uses Aspose.Cells WorkbookDesigner to replace an `Image:Photo` smart marker in cell A1 with the image. | Show how to configure a DataSet as the data source for WorkbookDesigner so that binary image data stored in a DataTable is rendered as an image marker in an Excel worksheet. | Demonstrate the complete workflow: create a workbook, place an `Image:{Property}` marker, load image bytes, bind them to a custom object, process the markers, and save the resulting .xlsx file.
// Common Searches: aspocells c# embed image in cell using smart markers from byte array | how to use WorkbookDesigner to replace Image:Photo marker with picture | load png into DataTable for Aspose.Cells image marker | Aspose.Cells smart image marker example with custom object source | C# generate Excel file with images from binary data using Aspose
// Tags: Aspose.Cells WorkbookDesigner image marker | embed PNG into Excel cell C# | binary image data source for smart markers | DataSet to Excel image marker Aspose | smart marker Image:Photo usage

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace ImageMarkerExample
{
    // Custom class that holds binary image data
    // The sample creates a new workbook, inserts an `Image:Photo` smart marker into cell A1, reads a PNG file into a byte array, stores the bytes in a DataTable within a DataSet, sets this DataSet as the data source for WorkbookDesigner, processes the marker to embed the image, and saves the result as ImageMarkerResult.xlsx.
    public class ImageSource
    {
        // Property name will be used in the marker (e.g., Image:Photo)
        public byte[]? Photo { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // 3. Place an image marker in a cell (A1). The marker format is "Image:{PropertyName}"
                sheet.Cells["A1"].PutValue("Image:Photo");

                // 4. Load binary image data from a file (or any other source)
                const string imagePath = "sample.png";
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }
                byte[] imageBytes = File.ReadAllBytes(imagePath);

                // 5. Create a DataSet with a DataTable that contains the image bytes
                DataTable table = new DataTable("ImageSource");
                table.Columns.Add("Photo", typeof(byte[]));
                table.Rows.Add(imageBytes);
                DataSet dataSet = new DataSet();
                dataSet.Tables.Add(table);

                // 6. Use WorkbookDesigner to process the image marker.
                //    The designer will replace the marker with the actual image.
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dataSet);
                designer.Process(); // Replaces the marker with the image

                // 7. Save the resulting workbook
                const string outputPath = "ImageMarkerResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
