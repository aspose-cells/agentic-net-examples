// Title: How to verify that embedded images are preserved when merging Excel workbooks using Workbook.Combine in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# combine two workbooks and keep pictures | check if pictures are retained after Workbook.Combine in .NET | merge Excel files with embedded images using Aspose.Cells | count pictures in combined workbook Aspose.Cells C# | preserve embedded PNGs when merging workbooks with Workbook.Combine
// Tags: Workbook.Combine image retention | Aspose.Cells merge workbooks with embedded pictures | C# verify picture count after combine | Excel workbook combine preserving graphics | temporary workbook picture verification Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example creates two temporary workbooks, adds a PNG picture to each worksheet, saves and reloads them, then merges the workbooks with Workbook.Combine. It iterates through all worksheets in the combined workbook, sums the Pictures.Count, and prints the total to confirm that both embedded images are retained before cleaning up the temporary files.
class Program
{
    static void Main()
    {
        // Create first workbook and embed an image
        Workbook wb1 = new Workbook();
        Worksheet ws1 = wb1.Worksheets[0];
        // Add a picture from a sample image file (ensure the file exists)
        string imagePath = "sample1.png"; // replace with actual image path
        if (File.Exists(imagePath))
        {
            ws1.Pictures.Add(1, 1, imagePath);
        }

        // Save first workbook to a temporary file
        string tempFile1 = Path.GetTempFileName();
        wb1.Save(tempFile1);

        // Create second workbook and embed another image
        Workbook wb2 = new Workbook();
        Worksheet ws2 = wb2.Worksheets[0];
        string imagePath2 = "sample2.png"; // replace with actual image path
        if (File.Exists(imagePath2))
        {
            ws2.Pictures.Add(2, 2, imagePath2);
        }

        // Save second workbook to a temporary file
        string tempFile2 = Path.GetTempFileName();
        wb2.Save(tempFile2);

        // Load the workbooks (using the provided load mechanism)
        Workbook loadWb1 = new Workbook(tempFile1);
        Workbook loadWb2 = new Workbook(tempFile2);

        // Combine the two workbooks into a new workbook
        Workbook combined = new Workbook();
        combined.Combine(loadWb1);
        combined.Combine(loadWb2);

        // Verify that embedded images are retained after combine
        int totalPictures = 0;
        foreach (Worksheet sheet in combined.Worksheets)
        {
            totalPictures += sheet.Pictures.Count;
        }

        Console.WriteLine($"Total pictures in combined workbook: {totalPictures}");
        // Expect totalPictures to be at least 2 if both images were retained

        // Clean up temporary files
        File.Delete(tempFile1);
        File.Delete(tempFile2);
    }
}
