using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerImageDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Place a smart marker tag that expects an image.
                // The tag name "ProductImage" will be used as the data source key.
                sheet.Cells["A1"].PutValue("&=ProductImage");

                // Load image data into a byte array (replace with your own image path).
                const string imagePath = "sample_image.png";
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                byte[] imageBytes = File.ReadAllBytes(imagePath);

                // ---------- Process smart markers ----------
                // Use WorkbookDesigner (the classic API) to replace the smart marker with the supplied image.
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("ProductImage", imageBytes);
                designer.Process();

                // ---------- Save the workbook ----------
                const string outputPath = "SmartMarkerImageOutput.xlsx";
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved successfully to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}