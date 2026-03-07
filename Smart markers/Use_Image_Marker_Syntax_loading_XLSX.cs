using System;
using System.Data;
using System.IO;
using Aspose.Cells;

class ImageMarkerDemo
{
    static void Main()
    {
        // Create a new workbook and add a marker for the image.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Employees";
        // Place the image marker in cell A1. The marker name must match the DataTable column name.
        sheet.Cells["A1"].PutValue("&Image=Photo");

        // Create a DataTable with a byte[] column for the image.
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Photo", typeof(byte[]));

        // Load an image file into a byte array.
        // Ensure the image file exists in the same directory as the executable or provide a full path.
        string imagePath = "photo.jpg";
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"Image file '{imagePath}' not found.");
            return;
        }
        byte[] photoBytes = File.ReadAllBytes(imagePath);
        dt.Rows.Add(photoBytes);

        // Initialize WorkbookDesigner, assign the workbook and the data source.
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };
        designer.SetDataSource(dt);

        // Process the smart markers. The image marker will be replaced with the image bytes.
        designer.Process();

        // Save the resulting workbook.
        workbook.Save("ResultWithImage.xlsx");
        Console.WriteLine("Workbook saved as 'ResultWithImage.xlsx'.");
    }
}