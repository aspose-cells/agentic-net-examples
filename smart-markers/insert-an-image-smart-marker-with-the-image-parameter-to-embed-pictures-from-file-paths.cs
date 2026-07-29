// Title: Insert an Image from a File Path Using Aspose.Cells Smart Markers (C#)
// Description: Demonstrates how to embed a picture in an Excel workbook by placing the smart‑marker "&=Image" in a cell, supplying a DataTable with an "Image" column that holds the full file path, processing the marker with WorkbookDesigner, optionally setting the picture to be placed inside the cell, and saving the result as an XLSX file.
// Keywords: Aspose.Cells smart marker image | C# embed picture from file path | WorkbookDesigner insert image | place picture inside cell Aspose.Cells | image smart marker syntax | Aspose.Cells .NET image insertion | Excel smart marker picture
// Common Searches: Aspose.Cells insert image using smart marker C# | How to embed picture from file path with Aspose.Cells | Smart marker syntax for images Aspose.Cells | Place picture inside a cell using WorkbookDesigner | C# Aspose.Cells image smart marker example
// Developer Intent: Embed an image into an Excel worksheet by using a smart marker that reads the image file path from a data source.
// Use Cases: Generate product catalogs where each product row automatically displays its photo from a stored image file. | Create employee directories that pull profile pictures from disk into the spreadsheet. | Automate report generation that inserts dynamically selected chart images stored on the server.
// AI Prompts: Show how to modify the code to handle multiple image paths and insert them into consecutive rows. | Provide an example that uses relative image paths and automatically resizes pictures to fit cell dimensions. | Explain how to replace the DataTable with a List<T> as the data source for image smart markers.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartMarkerImageDemo
{
    // Demonstrates how to embed a picture in an Excel workbook by placing the smart‑marker "&=Image" in a cell, supplying a DataTable with an "Image" column that holds the full file path, processing the marker with WorkbookDesigner, optionally setting the picture to be placed inside the cell, and saving the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Insert a smart marker that expects an image path.
                // The marker syntax "&=Image" tells Aspose.Cells to treat the value as an image.
                sheet.Cells["A1"].PutValue("&=Image");

                // Prepare data source: a DataTable with a column named "Image"
                // containing the full file path of the picture to embed.
                DataTable dt = new DataTable("Images");
                dt.Columns.Add("Image", typeof(string));

                // Define the image file path (adjust to an existing image on your machine).
                string imagePath = @"C:\Images\sample_picture.jpg";

                // Ensure the image file exists before adding it to the data source.
                if (File.Exists(imagePath))
                {
                    dt.Rows.Add(imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    // Add an empty string to keep the row count consistent.
                    dt.Rows.Add(string.Empty);
                }

                // Process the smart markers using WorkbookDesigner.
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);
                designer.Process();

                // Optionally, adjust the picture placement to fit inside the cell.
                // Retrieve the inserted picture (it will be the last picture added).
                if (sheet.Pictures.Count > 0)
                {
                    Picture pic = sheet.Pictures[sheet.Pictures.Count - 1];
                    pic.IsPlacedInCell = true; // place picture inside the cell
                }

                // Save the workbook.
                string outputPath = "SmartMarkerImageOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
