// Title: Insert an Image Smart Marker from a File Path using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to place an image smart marker ("&=\"Image\"") in a worksheet, read a JPEG file into a byte array, bind it via a DataSet, process with WorkbookDesigner, and save the result as an Excel file.
// Keywords: Aspose.Cells | C# | image smart marker | embed picture Excel | WorkbookDesigner | byte[] image | file path image | Excel automation | smart markers tutorial | .NET Excel library
// Common Searches: Aspose.Cells insert image from file path | C# smart marker image example | WorkbookDesigner embed picture Excel | how to use image smart markers Aspose | load JPEG into Excel using Aspose.Cells
// Developer Intent: Replace a smart marker with a picture loaded from a local file and generate the final workbook programmatically.
// Use Cases: Create a product catalog where each item’s photo is inserted automatically via a smart marker. | Generate an employee directory that places staff headshots into the spreadsheet without manual editing. | Add a company logo to report headers by processing a single image smart marker during workbook creation.
// AI Prompts: Show how to modify the sample to insert multiple images from a DataTable with several rows of byte[] values. | Provide code that logs missing image files but continues processing other smart markers. | Explain how to control the inserted image’s dimensions or scaling when using an image smart marker.

using System;
using System.IO;
using System.Data;
using Aspose.Cells;

// Demonstrates how to place an image smart marker ("&=\"Image\"") in a worksheet, read a JPEG file into a byte array, bind it via a DataSet, process with WorkbookDesigner, and save the result as an Excel file.
public class ImageData
{
    // Property name matches the smart marker name; using byte[] avoids System.Drawing dependency
    public byte[] Image { get; set; } = Array.Empty<byte>();
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook (template)
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Insert an image smart marker in cell A1
            // The marker syntax &="Image" tells Aspose.Cells to replace it with the Image property value
            worksheet.Cells["A1"].PutValue("&=\"Image\"");

            // Prepare the data source containing the image to embed
            string imagePath = "sample.jpg"; // path to the image file
            var data = new ImageData();

            try
            {
                if (File.Exists(imagePath))
                {
                    data.Image = File.ReadAllBytes(imagePath);
                }
                else
                {
                    Console.WriteLine($"Warning: Image file '{imagePath}' not found. An empty image will be used.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading image file: {ex.Message}");
                data.Image = Array.Empty<byte>();
            }

            // Build a DataSet with a DataTable that matches the smart marker name
            var dataTable = new DataTable("ImageData");
            dataTable.Columns.Add("Image", typeof(byte[]));
            dataTable.Rows.Add(data.Image);
            var dataSet = new DataSet();
            dataSet.Tables.Add(dataTable);

            // Process the smart marker and embed the image using WorkbookDesigner
            var designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // Save the resulting workbook
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
