// Title: How to embed a Base64‑encoded image into an Excel worksheet using Aspose.Cells smart image markers in C#
// Common Searches: C# Aspose.Cells replace smart marker with Base64 image in Excel | embed Base64 photo into Excel using WorkbookDesigner and image marker | Aspose.Cells smart image marker from DataTable Base64 string example | load image from Base64 string into Excel cell with Aspose.Cells C# | use ${photo} placeholder to insert decoded picture in an Excel template
// Tags: smart image marker Aspose.Cells | Base64 image embedding Excel C# | WorkbookDesigner SetDataSource DataTable | decode Base64 string to worksheet picture | image placeholder ${photo} replacement

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, places a smart image marker ${photo} in cell A1, builds a DataTable containing a Base64‑encoded image read from a file, sets the table as the data source for a WorkbookDesigner, processes the markers so the placeholder is replaced with the decoded picture, and saves the result as output.xlsx.
class EmbedBase64Images
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Template";

        // Place an image marker in cell A1. The marker name must match the column name in the data source.
        sheet.Cells["A1"].PutValue("${photo}");

        // Prepare the data source with a Base64 string representing an image
        DataTable dt = new DataTable();
        dt.Columns.Add("photo", typeof(string));

        // Example: read an image file, convert it to Base64 string, and add to the table
        // Replace "sample.jpg" with the path to your image file
        string imagePath = "sample.jpg";
        if (File.Exists(imagePath))
        {
            byte[] imageBytes = File.ReadAllBytes(imagePath);
            string base64String = Convert.ToBase64String(imageBytes);
            dt.Rows.Add(base64String);
        }
        else
        {
            // If the image file does not exist, add an empty string to avoid errors
            dt.Rows.Add(string.Empty);
        }

        // Use WorkbookDesigner to process smart markers
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);
        designer.Process(); // This replaces the marker with the decoded image

        // Save the resulting workbook
        workbook.Save("output.xlsx");
    }
}
