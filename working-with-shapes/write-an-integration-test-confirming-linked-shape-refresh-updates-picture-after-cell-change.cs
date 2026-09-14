// Title: C# integration test to verify that a linked picture shape refreshes after its linked cell value changes using Aspose.Cells
// AI Prompts: Generate a C# integration test that adds a PNG picture to a worksheet, links it to cell A1, calls RefreshLinkedShape before and after modifying the cell value, and asserts that the PictureFormat.ImageBytes arrays are different. | Rewrite the test to use a JPEG file instead of PNG and confirm the linked shape updates correctly when the cell content is changed. | Add error handling so that if IsLinked or RefreshLinkedShape are unavailable in the current Aspose.Cells version, the test logs the limitation and still guarantees cleanup of the temporary image file.
// Common Searches: Aspose.Cells verify picture updates after linked cell change in C# | how to unit test RefreshLinkedShape method with a temporary image | C# example linking a picture to a cell and refreshing it in Aspose.Cells | compare picture byte arrays before and after cell value modification Aspose.Cells | integration test for picture shape linked to cell A1 using Aspose.Cells .NET
// Tags: picture shape refresh verification | temporary PNG image for shape linking | image byte array comparison after cell edit | Aspose.Cells linked cell picture test | handling unsupported picture linking

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Creates a temporary 1x1 PNG, inserts it as a picture linked to cell A1, refreshes the linked shape before and after changing the cell value, captures the image bytes each time, asserts the byte arrays differ, and cleans up the temporary file.
    class Program
    {
        static void Main()
        {
            // Create a temporary PNG image file (1x1 pixel) to serve as the source picture.
            string tempImagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
            try
            {
                // PNG byte array for a 1x1 transparent pixel.
                byte[] pngBytes = Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
                File.WriteAllBytes(tempImagePath, pngBytes);

                // Verify the temporary image file exists.
                if (!File.Exists(tempImagePath))
                {
                    Console.WriteLine("Failed to create temporary image file.");
                    return;
                }

                // Create a new workbook and obtain the first worksheet.
                var workbook = new Workbook();
                var ws = workbook.Worksheets[0];

                // Add a picture to the worksheet.
                int pictureIndex = ws.Pictures.Add(0, 0, 0, 0, tempImagePath);
                dynamic picture = ws.Pictures[pictureIndex]; // Use dynamic to access members that may vary between versions.

                try
                {
                    // Link the picture to cell A1 (if supported by the current Aspose.Cells version).
                    picture.IsLinked = true;
                    picture.LinkedCell = "A1";

                    // Set the initial value of the linked cell and refresh the picture.
                    ws.Cells["A1"].PutValue("Initial");
                    picture.RefreshLinkedShape();

                    // Capture the picture bytes after the first refresh.
                    byte[] bytesBefore = picture.PictureFormat.ImageBytes;

                    // Change the cell value and refresh the picture again.
                    ws.Cells["A1"].PutValue("Updated");
                    picture.RefreshLinkedShape();

                    // Capture the picture bytes after the second refresh.
                    byte[] bytesAfter = picture.PictureFormat.ImageBytes;

                    // Verify that the picture was updated (the image bytes differ).
                    bool success = true;

                    if (bytesBefore == null || bytesBefore.Length == 0)
                    {
                        Console.WriteLine("Initial picture bytes should not be empty.");
                        success = false;
                    }

                    if (bytesAfter == null || bytesAfter.Length == 0)
                    {
                        Console.WriteLine("Updated picture bytes should not be empty.");
                        success = false;
                    }

                    if (bytesBefore != null && bytesAfter != null && AreArraysEqual(bytesBefore, bytesAfter))
                    {
                        Console.WriteLine("Picture should be refreshed and differ after cell change.");
                        success = false;
                    }

                    Console.WriteLine(success ? "Test passed." : "Test failed.");
                }
                catch (Exception picEx)
                {
                    // If linking or refresh is not supported, report and continue.
                    Console.WriteLine($"Picture operation exception: {picEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An exception occurred: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary image file.
                try
                {
                    if (File.Exists(tempImagePath))
                    {
                        File.Delete(tempImagePath);
                    }
                }
                catch
                {
                    // Suppress any exceptions during cleanup.
                }
            }
        }

        private static bool AreArraysEqual(byte[] a, byte[] b)
        {
            if (a == null || b == null) return false;
            if (a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i]) return false;
            }
            return true;
        }
    }
}
