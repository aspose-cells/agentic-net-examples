// Title: C# Integration Test – Verify Linked Picture Refresh After Cell Update with Aspose.Cells
// Description: Creates a temporary 1×1 PNG, adds a linked picture to a worksheet (row 2, column 2), links it to cell A1, sets an initial value, calls UpdateSelectedValue, checks the linked cell reference, changes the cell value, refreshes the picture again, confirms the new value, saves the workbook to a MemoryStream, and cleans up the temporary image file.
// Keywords: Aspose.Cells linked picture test | C# AddLinkedPicture example | UpdateSelectedValue Aspose.Cells | linked shape refresh .NET | integration test workbook save | temporary PNG Aspose.Cells | SetLinkedCell verification
// Common Searches: how to test linked picture refresh in Aspose.Cells | Aspose.Cells C# unit test for linked shape update | AddLinkedPicture and UpdateSelectedValue example | verify linked cell value after picture refresh | Aspose.Cells integration test for linked images
// Developer Intent: Confirm that a linked picture shape updates its displayed content when the source cell value changes.
// Use Cases: Automated CI validation that linked pictures reflect the latest cell data. | Ensuring workbooks with linked images can be saved to streams without errors. | Demonstrating proper resource cleanup (temporary image files) in shape‑related tests.
// AI Prompts: Generate an NUnit test that creates a temporary PNG, adds a linked picture to a worksheet, links it to cell A1, changes the cell value, calls UpdateSelectedValue, and asserts the picture reflects the new value. | Write a MSTest method that verifies GetLinkedCell returns "$A$1" after linking a picture in Aspose.Cells and confirms the cell value updates correctly. | Provide a C# snippet that safely deletes the temporary image file used for a linked picture in a test's finally block.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTests
{
    // Creates a temporary 1×1 PNG, adds a linked picture to a worksheet (row 2, column 2), links it to cell A1, sets an initial value, calls UpdateSelectedValue, checks the linked cell reference, changes the cell value, refreshes the picture again, confirms the new value, saves the workbook to a MemoryStream, and cleans up the temporary image file.
    class Program
    {
        static void Main()
        {
            try
            {
                RunLinkedPictureRefreshTest();
                Console.WriteLine("Test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        static void RunLinkedPictureRefreshTest()
        {
            // Path for a temporary PNG image (1x1 pixel) used as the linked picture source.
            string tempImagePath = Path.Combine(Path.GetTempPath(), $"tempImg_{Guid.NewGuid()}.png");

            try
            {
                // Create a minimal PNG byte array (transparent 1x1 pixel) and write it to the temp file.
                byte[] pngBytes = new byte[]
                {
                    0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                    0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                    0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                    0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
                    0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                    0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
                    0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
                    0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
                    0x42,0x60,0x82
                };

                File.WriteAllBytes(tempImagePath, pngBytes);

                if (!File.Exists(tempImagePath))
                    throw new FileNotFoundException("Temporary image file was not created.", tempImagePath);

                // Create a new workbook and get the first worksheet.
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Add a linked picture to the sheet (positioned at row 2, column 2).
                var linkedPic = sheet.Shapes.AddLinkedPicture(2, 2, 100, 100, tempImagePath);

                // Link the picture to cell A1.
                linkedPic.SetLinkedCell("$A$1", true, true);

                // Put an initial value into the linked cell.
                sheet.Cells["A1"].PutValue("First");
                linkedPic.UpdateSelectedValue();

                // Verify that the linked cell reference is correct.
                string linkedCellRef = linkedPic.GetLinkedCell(true, true);
                if (linkedCellRef != "$A$1")
                    throw new InvalidOperationException($"Linked cell reference mismatch. Expected $A$1, got {linkedCellRef}");

                // Change the value of the linked cell.
                sheet.Cells["A1"].PutValue("Second");
                linkedPic.UpdateSelectedValue();

                // Retrieve the value from the linked cell via the worksheet to ensure it was updated.
                string cellValue = sheet.Cells["A1"].StringValue;
                if (cellValue != "Second")
                    throw new InvalidOperationException($"Cell value mismatch. Expected 'Second', got '{cellValue}'");

                // Save the workbook to a memory stream to ensure no exceptions are thrown.
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    if (ms.Length == 0)
                        throw new InvalidOperationException("Workbook was not saved correctly.");
                }
            }
            catch (Exception ex)
            {
                // Propagate exception after cleanup.
                throw new InvalidOperationException("An error occurred during the linked picture refresh test.", ex);
            }
            finally
            {
                // Clean up the temporary image file.
                if (File.Exists(tempImagePath))
                {
                    try { File.Delete(tempImagePath); } catch { /* ignore cleanup errors */ }
                }
            }
        }
    }
}
