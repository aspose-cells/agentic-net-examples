// Title: Insert Image via Smart Marker Using a Byte[] Data Source – Aspose.Cells C# Example
// Description: Demonstrates how to place an image in an Excel worksheet by adding a smart marker "&Image:ImageData" to a cell, supplying a DataTable with a byte[] column, and processing the marker with WorkbookDesigner. The example loads a PNG file into a byte array, binds it as the data source, and saves the result as SmartMarkerImageOutput.xlsx.
// Keywords: Aspose.Cells | C# smart marker image | byte array image Excel | WorkbookDesigner | Insert image smart marker | Excel image from database | Aspose.Cells example | GitHub Aspose.Cells image smart marker
// Common Searches: Aspose.Cells insert image from byte array C# | smart marker &Image usage example | WorkbookDesigner image data source tutorial | C# code to embed PNG in Excel with Aspose.Cells | GitHub repository Aspose.Cells smart marker image
// Developer Intent: Add an image to an Excel file by using a smart marker that references a byte[] field.
// Use Cases: Generate product catalogs where each item’s photo is stored as a BLOB and inserted via smart markers. | Create employee directories that display profile pictures retrieved from a database. | Automate reports that embed chart screenshots saved as byte arrays into the final workbook.
// AI Prompts: Write a C# program that reads all PNG files from a folder, fills a DataTable with byte[] values, and uses Aspose.Cells smart markers to insert each image into separate rows of an Excel sheet. | Explain how to control image size, scaling, and alignment when using the &Image smart marker in Aspose.Cells. | Provide error‑handling code for missing or corrupted image files while processing smart markers with WorkbookDesigner.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// Demonstrates how to place an image in an Excel worksheet by adding a smart marker "&Image:ImageData" to a cell, supplying a DataTable with a byte[] column, and processing the marker with WorkbookDesigner. The example loads a PNG file into a byte array, binds it as the data source, and saves the result as SmartMarkerImageOutput.xlsx.
class InsertImageSmartMarker
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a smart marker that will be replaced by an image.
            // The marker expects a byte[] field named ImageData.
            sheet.Cells["A1"].PutValue("&Image:ImageData");

            // Prepare a data source containing the image bytes.
            DataTable dt = new DataTable("Images");
            dt.Columns.Add("ImageData", typeof(byte[]));

            // Load an image file into a byte array.
            string imagePath = "example.png";
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            byte[] imageBytes = File.ReadAllBytes(imagePath);
            dt.Rows.Add(imageBytes);

            // Process the smart markers using WorkbookDesigner.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process();

            // Save the workbook with the inserted image.
            string outputPath = "SmartMarkerImageOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
