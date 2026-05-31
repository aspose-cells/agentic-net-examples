using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTests
{
    class Program
    {
        static void Main()
        {
            try
            {
                var test = new LinkedPictureRefreshTests();
                test.GlobalSetup();
                test.LinkedPicture_ShouldRefresh_WhenLinkedCellValueChanges();
                test.GlobalTeardown();
                Console.WriteLine("Test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class LinkedPictureRefreshTests
    {
        private string? _imagePath1;
        private string? _imagePath2;
        private string? _tempFolder;

        // Small 1x1 PNG images (red and green) encoded in Base64
        private const string RedPngBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
        private const string GreenPngBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/6ZcVAAAAAElFTkSuQmCC";

        public void GlobalSetup()
        {
            try
            {
                // Create temporary folder for test images
                _tempFolder = Path.Combine(Path.GetTempPath(), "AsposeLinkedPictureTest");
                Directory.CreateDirectory(_tempFolder);

                // Red square image
                _imagePath1 = Path.Combine(_tempFolder, "red.png");
                File.WriteAllBytes(_imagePath1, Convert.FromBase64String(RedPngBase64));

                // Green square image
                _imagePath2 = Path.Combine(_tempFolder, "green.png");
                File.WriteAllBytes(_imagePath2, Convert.FromBase64String(GreenPngBase64));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GlobalSetup failed: {ex.Message}");
                throw;
            }
        }

        public void GlobalTeardown()
        {
            try
            {
                // Delete temporary folder and its contents
                if (!string.IsNullOrEmpty(_tempFolder) && Directory.Exists(_tempFolder))
                {
                    Directory.Delete(_tempFolder, true);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GlobalTeardown warning: {ex.Message}");
            }
        }

        public void LinkedPicture_ShouldRefresh_WhenLinkedCellValueChanges()
        {
            try
            {
                // Ensure test images exist
                if (string.IsNullOrEmpty(_imagePath1) || string.IsNullOrEmpty(_imagePath2) ||
                    !File.Exists(_imagePath1) || !File.Exists(_imagePath2))
                {
                    throw new FileNotFoundException("Test images not found.");
                }

                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];

                // Put the initial image path into cell A1
                worksheet.Cells["A1"].PutValue(_imagePath1);

                // Add a linked picture that initially points to the image in A1
                var picture = worksheet.Shapes.AddLinkedPicture(2, 2, 100, 100, _imagePath1);

                // Link the picture to cell A1 so it refreshes when the cell changes
                picture.SetLinkedCell("$A$1", false, false);

                // Verify the linked cell is set correctly
                if (picture.GetLinkedCell(false, false) != "$A$1")
                    throw new InvalidOperationException("Linked cell not set correctly.");

                // Change the cell value to point to a different image
                worksheet.Cells["A1"].PutValue(_imagePath2);

                // Refresh the picture so it reads the new path from the linked cell
                picture.UpdateSelectedValue();

                // Verify the picture still exists in the collection after refresh
                if (!worksheet.Shapes.Contains(picture))
                    throw new InvalidOperationException("Picture was removed after update.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
                throw;
            }
        }
    }
}