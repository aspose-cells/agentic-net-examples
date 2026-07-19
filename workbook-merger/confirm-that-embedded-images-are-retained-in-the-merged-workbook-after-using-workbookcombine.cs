// Title: Aspose.Cells C# – Verify Embedded Images Remain After Workbook.Combine
// Description: Demonstrates how to embed a PNG into a cell, merge the source workbook into a destination workbook with Workbook.Combine, count embedded pictures before and after the merge, and save the result to confirm that the image is preserved.
// Keywords: Aspose.Cells Workbook.Combine | C# embedded image in cell | preserve cell pictures after merge | CountPlaceInCellPictures Aspose | Combine workbooks with images | EmbeddedImage property | .NET Excel merge image retention
// Common Searches: keep embedded pictures when using Workbook.Combine | does Workbook.Combine copy cell images | Aspose.Cells count embedded images after merge | sample code to test image retention in combined workbook | C# verify embedded image after workbook combine
// Developer Intent: Confirm that cell‑embedded images survive the Workbook.Combine operation and can be programmatically validated.
// Use Cases: Quality‑check merged financial reports that contain company logos. | Automate consolidation of multiple worksheets while retaining in‑cell graphics. | Validate that PlaceInCell pictures are not lost during batch workbook merging.
// AI Prompts: Generate C# code using Aspose.Cells to merge two workbooks and assert that the number of embedded images in the result equals the source count. | Explain how Workbook.Combine processes PlaceInCell pictures and which cell properties indicate successful preservation. | Suggest alternative merging techniques (e.g., copying worksheets) that guarantee embedded image retention.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// Demonstrates how to embed a PNG into a cell, merge the source workbook into a destination workbook with Workbook.Combine, count embedded pictures before and after the merge, and save the result to confirm that the image is preserved.
class EmbeddedImageCombineDemo
{
    static void Main()
    {
        try
        {
            // Path to a sample image file (ensure this file exists on disk)
            string imagePath = "sample.png";

            // Verify that the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // -------------------- Create source workbook --------------------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Embed the image into cell B2 of the source workbook
            sourceSheet.Cells["B2"].EmbeddedImage = File.ReadAllBytes(imagePath);

            // Verify that the source workbook contains one embedded image
            int sourceImageCount = CountEmbeddedImages(sourceSheet.Cells);
            Console.WriteLine($"Source workbook embedded images: {sourceImageCount}");

            // -------------------- Create destination workbook --------------------
            Workbook destinationWorkbook = new Workbook();

            // -------------------- Combine workbooks --------------------
            destinationWorkbook.Combine(sourceWorkbook);

            // After combining, verify that the embedded image is retained
            Worksheet destSheet = destinationWorkbook.Worksheets[0];
            int destImageCount = CountEmbeddedImages(destSheet.Cells);
            Console.WriteLine($"Destination workbook after combine embedded images: {destImageCount}");

            // Save the combined workbook for manual inspection
            destinationWorkbook.Save("CombinedWithImages.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to count cells that contain embedded pictures (PlaceInCell)
    static int CountEmbeddedImages(Cells cells)
    {
        int count = 0;
        IEnumerator enumerator = cells.GetCellsWithPlaceInCellPicture();
        if (enumerator != null)
        {
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                if (cell.EmbeddedImage != null && cell.EmbeddedImage.Length > 0)
                {
                    count++;
                }
            }
        }
        return count;
    }
}
