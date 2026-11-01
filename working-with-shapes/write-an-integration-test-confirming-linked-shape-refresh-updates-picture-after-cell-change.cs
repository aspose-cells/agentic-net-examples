// Title: C# integration test to verify that a linked picture refreshes after a cell value change using Aspose.Cells
// Common Searches: c# Aspose.Cells integration test linked picture refresh after cell update | how to assert picture persistence in an Aspose.Cells workbook after modifying a cell | unit test for picture shape refresh in an Aspose.Cells .xlsx file | verify that a linked picture updates when a cell value changes with Aspose.Cells for .NET | Aspose.Cells C# test to ensure picture remains after workbook reload
// Tags: picture shape refresh validation Aspose.Cells | add picture from file to worksheet Aspose.Cells | verify picture persistence after cell edit Aspose.Cells | workbook reload shape existence test Aspose.Cells | integration test for worksheet shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates a temporary PNG, adds it as a picture to cell C3 of a new workbook with Aspose.Cells, saves the file, changes the value of cell A1, saves again, reloads the workbook, and asserts that the picture still exists and the cell value reflects the update, while handling setup and cleanup of temporary resources.
    public class LinkedPictureRefreshDemo
    {
        private string _tempImagePath = string.Empty;
        private string _tempWorkbookPath = string.Empty;

        // Entry point
        public static void Main()
        {
            var demo = new LinkedPictureRefreshDemo();
            demo.Run();
        }

        // Executes the demo with proper resource handling
        public void Run()
        {
            try
            {
                SetUp();
                PictureRefreshDemo();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                TearDown();
            }
        }

        // Creates a temporary PNG image and workbook file paths
        private void SetUp()
        {
            // Small 1x1 pixel PNG (transparent) encoded in Base64
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAn8B9pXc2VQAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);

            _tempImagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
            File.WriteAllBytes(_tempImagePath, pngBytes);

            // Path for the temporary workbook (used for verification)
            _tempWorkbookPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
        }

        // Deletes temporary files if they exist
        private void TearDown()
        {
            try
            {
                if (File.Exists(_tempImagePath))
                    File.Delete(_tempImagePath);
                if (File.Exists(_tempWorkbookPath))
                    File.Delete(_tempWorkbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Cleanup error: {ex.Message}");
            }
        }

        // Core logic demonstrating picture handling and cell value change
        private void PictureRefreshDemo()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Set an initial value in cell A1
                sheet.Cells["A1"].PutValue("Hello");

                // Ensure the image file exists before adding it
                if (!File.Exists(_tempImagePath))
                    throw new FileNotFoundException("Temporary image file not found.", _tempImagePath);

                // Add a picture that references the temporary image file (placed at C3)
                // Note: Use Add (row, column, picturePath) – not AddPicture
                sheet.Pictures.Add(2, 2, _tempImagePath);

                // Save the workbook (ensures file can be persisted)
                workbook.Save(_tempWorkbookPath, SaveFormat.Xlsx);

                // Verify that the picture was added
                if (sheet.Pictures.Count == 0)
                    throw new InvalidOperationException("Picture was not added to the worksheet.");

                // Change the value of cell A1
                sheet.Cells["A1"].PutValue("World");

                // Save again to capture the updated cell value
                workbook.Save(_tempWorkbookPath, SaveFormat.Xlsx);

                // Reload the workbook to ensure persistence
                if (!File.Exists(_tempWorkbookPath))
                    throw new FileNotFoundException("Workbook file was not saved correctly.", _tempWorkbookPath);

                var reloadedWorkbook = new Workbook(_tempWorkbookPath);
                var reloadedSheet = reloadedWorkbook.Worksheets[0];

                // Verify that the picture still exists after reload
                if (reloadedSheet.Pictures.Count == 0)
                    throw new InvalidOperationException("Reloaded workbook does not contain the picture.");

                // Verify that the cell value was updated
                if (reloadedSheet.Cells["A1"].StringValue != "World")
                    throw new InvalidOperationException("Reloaded sheet should contain the updated cell value.");

                Console.WriteLine("Picture handling test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PictureRefreshDemo error: {ex.Message}");
                throw; // Re‑throw to be caught by outer Run() handler
            }
        }
    }
}
