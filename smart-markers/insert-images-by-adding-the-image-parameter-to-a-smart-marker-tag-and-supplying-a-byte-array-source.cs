// Title: Insert an image into an Excel worksheet using Aspose.Cells smart markers with a byte[] data source (C#)
// Common Searches: how to embed a picture in Excel using Aspose.Cells smart markers C# | using a byte array as image source for Aspose.Cells WorkbookDesigner | populate smart marker {Image} with image bytes in C# | Aspose.Cells insert image from DataTable column byte[] | C# generate Excel file with embedded image via smart markers
// Tags: smart marker image insertion Aspose.Cells | WorkbookDesigner embed picture using image bytes | DataTable column for image bytes Excel | C# generate Excel with embedded image via smart markers | Aspose.Cells load image bytes into worksheet

using Aspose.Cells;
using System;
using System.Data;
using System.IO;

// The example creates a workbook, places a smart marker tag {Image} in cell A1, reads an image file into a byte array, adds the byte array to a DataTable column named "Image", sets this table as the data source for a WorkbookDesigner, processes the smart marker to embed the picture, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a smart marker that expects an image parameter.
            sheet.Cells["A1"].PutValue("{Image}");

            // Prepare a data source with a byte[] column for the image
            DataTable dt = new DataTable();
            dt.Columns.Add("Image", typeof(byte[]));

            // Load image bytes safely
            string imagePath = "sample.png";
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            byte[] imgBytes = File.ReadAllBytes(imagePath);
            dt.Rows.Add(imgBytes);

            // Process the smart marker using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            // KeepSmartMarkers property is not available in this version; omitted.
            designer.Process();

            // Ensure output directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook to a file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
