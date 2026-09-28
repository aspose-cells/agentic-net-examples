// Title: Insert a PNG image into an Excel workbook using Aspose.Cells smart markers and a byte[] data source in C#
// AI Prompts: Generate C# code that reads a PNG file into a memory buffer, places a '&=Image' smart marker in a worksheet cell, and uses WorkbookDesigner to embed the picture into an .xlsx file. | Show how to construct a DataSet that includes an image field containing binary data and process it with Aspose.Cells smart markers to replace the marker with the actual image. | Provide error‑handling steps for missing image files and demonstrate saving the workbook after the smart marker processing.
// Common Searches: how to embed a PNG into Excel using Aspose.Cells smart markers C# | Aspose.Cells WorkbookDesigner replace &=Image marker with picture data | C# create DataSet with image field for smart marker processing | smart marker example inserting images from memory in Aspose.Cells
// Tags: Aspose.Cells smart marker image embedding | C# load PNG for Excel insertion | WorkbookDesigner image marker processing | DataSet binary image for Aspose.Cells | embed PNG into .xlsx via smart markers

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example reads a PNG file into a byte array, places a '&=Image' smart marker in cell A1, builds a DataSet with an image column holding the byte data, processes the marker with WorkbookDesigner, and saves the result as SmartMarkerImage.xlsx.
class InsertImageWithSmartMarker
{
    static void Main()
    {
        try
        {
            // Verify that the image file exists before attempting to read it
            const string imagePath = "sample.png";
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Error: Image file '{imagePath}' not found.");
                return;
            }

            // Load the image file into a byte array (will be supplied to the smart marker)
            byte[] imageBytes = File.ReadAllBytes(imagePath);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Place a smart marker that expects an image (byte array) in cell A1
            // The marker syntax "&=Image" tells Aspose.Cells to replace it with image data
            worksheet.Cells["A1"].PutValue("&=Image");

            // Prepare the data source for the smart marker using a DataSet
            DataTable table = new DataTable("Data");
            table.Columns.Add("Image", typeof(byte[]));
            DataRow row = table.NewRow();
            row["Image"] = imageBytes;
            table.Rows.Add(row);

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(table);

            // Process the smart marker using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // Save the workbook with the inserted image
            const string outputPath = "SmartMarkerImage.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
