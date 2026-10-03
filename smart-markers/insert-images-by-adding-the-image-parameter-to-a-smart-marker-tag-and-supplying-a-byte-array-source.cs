// Title: Insert an image into an Excel worksheet using Aspose.Cells smart markers with a byte[] data source (C#)
// AI Prompts: Write C# code that loads a PNG file into a byte array, adds it to a DataTable, and uses WorkbookDesigner to replace a smart marker '&=Image' with the picture in an Excel file. | Show how to bind a DataTable containing a byte[] column to Aspose.Cells WorkbookDesigner and process smart markers to embed images. | Provide a complete example that creates a workbook, places a smart marker for an image, supplies the image bytes, processes the marker, and saves the result as an .xlsx file.
// Common Searches: Aspose.Cells C# smart marker image from byte array example | How to embed a PNG into Excel using WorkbookDesigner and smart markers | Insert pictures into Excel cells with Aspose.Cells smart markers and DataTable | C# load image as byte[] and use it as smart marker source in Aspose.Cells
// Tags: Aspose.Cells WorkbookDesigner image smart marker | C# byte[] image data source for Excel | smart marker picture insertion Excel | load PNG into byte array Aspose.Cells | process smart markers with image data

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, adds a smart marker '&=Image' to cell A1, reads a PNG file into a byte array, stores the bytes in a DataTable column named 'Image', binds the table to a WorkbookDesigner, processes the smart marker to embed the picture, and saves the workbook as OutputWithImage.xlsx.
class InsertImageWithSmartMarker
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a smart marker tag that expects an image.
            // The tag name "Image" will be matched with a column named "Image" in the data source.
            sheet.Cells["A1"].PutValue("&=Image");

            // Prepare a DataTable as the data source for the smart marker.
            DataTable dt = new DataTable();
            // The column must be of type byte[] to hold the image data.
            dt.Columns.Add("Image", typeof(byte[]));

            // Load an image file into a byte array.
            string imagePath = @"C:\Images\sample.png";
            if (!File.Exists(imagePath))
                throw new FileNotFoundException("Image file not found.", imagePath);

            byte[] imageBytes = File.ReadAllBytes(imagePath);

            // Add a row containing the image byte array.
            DataRow row = dt.NewRow();
            row["Image"] = imageBytes;
            dt.Rows.Add(row);

            // Process the smart markers using WorkbookDesigner (the correct API for smart markers).
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process();

            // Save the resulting workbook.
            string outputPath = "OutputWithImage.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
