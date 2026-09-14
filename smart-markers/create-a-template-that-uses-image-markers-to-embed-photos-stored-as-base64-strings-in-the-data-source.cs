// Title: Insert Base64‑encoded photos with Aspose.Cells smart image markers and export to HTML in C#
// AI Prompts: Write C# code that adds a smart image marker "&=$Photo" to a worksheet cell, binds a DataTable column containing a Base64 string, processes the marker with WorkbookDesigner, and saves the workbook as HTML with images embedded. | Show how to read an image file (e.g., PNG or JPEG), convert it to a Base64 string, place the string into a DataTable, and use Aspose.Cells to replace a cell with that image via a smart marker. | Modify the sample so the HTML output is written to a MemoryStream and returned as a byte array instead of being saved to a physical file.
// Common Searches: asp.net c# embed base64 image using aspose.cells smart marker | export aspose.cells workbook to html with images embedded as base64 | bind datatable column containing base64 string to aspose.cells image marker | use &=$Photo marker in aspose.cells to insert images from a database | convert image to base64 and display in excel using aspose.cells smart markers
// Tags: Aspose.Cells image marker with Base64 data | Set DataTable as data source for smart markers | HTML export with embedded Base64 images | Convert image file to Base64 string in C# | Replace worksheet cell with image using &=$Photo syntax

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsImageMarkerDemo
{
    // The example creates a workbook, places the smart image marker "&=$Photo" in cell A1, loads an image file, converts it to a Base64 string, adds the string to a DataTable, assigns the table as the data source for a WorkbookDesigner, processes the marker, and saves the workbook as HTML with the image embedded as a Base64 string.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Insert an image marker in cell A1.
            // The marker syntax "&=$Photo" tells the designer to replace the cell with an image
            // whose data comes from the "Photo" field of the data source.
            sheet.Cells["A1"].PutValue("&=$Photo");

            // 3. Prepare a DataTable that will serve as the data source.
            // The "Photo" column holds a Base64 string representation of an image.
            DataTable dt = new DataTable("Images");
            dt.Columns.Add("Photo", typeof(string));

            // 4. Load an image file, convert it to a Base64 string and add it to the table.
            // Replace "sample.jpg" with the path to your image file.
            string imagePath = "sample.jpg";
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            byte[] imageBytes = File.ReadAllBytes(imagePath);
            string base64Image = Convert.ToBase64String(imageBytes);
            dt.Rows.Add(base64Image);

            // 5. Create a WorkbookDesigner, assign the workbook and set the data source.
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;
            designer.SetDataSource(dt);

            // 6. Process the smart markers – the image marker will be replaced with the image.
            designer.Process();

            // 7. Save the result as HTML with images embedded as Base64 strings.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true   // Embed images directly in the HTML
            };

            string outputPath = "ImageMarkerOutput.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to {outputPath} with embedded Base64 images.");
        }
    }
}
