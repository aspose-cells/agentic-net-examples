// Title: How to verify that embedded images are preserved when merging two Excel workbooks using Workbook.Combine in Aspose.Cells for .NET
// AI Prompts: Create two workbooks, add a PNG picture to each worksheet, merge them with Workbook.Combine, then programmatically count the pictures to confirm both images remain. | Write C# code that loads two Excel files containing pictures, combines them via Aspose.Cells, validates that the images are retained, and saves the resulting workbook.
// Common Searches: Aspose.Cells Workbook.Combine keep pictures in merged file | C# merge Excel workbooks with embedded images using Aspose.Cells | count pictures after combining workbooks Aspose.Cells .NET | verify image retention after Workbook.Combine operation | preserve PNG images when merging Excel workbooks with Aspose.Cells
// Tags: Workbook.Combine preserve embedded images | Aspose.Cells count pictures after merge | merge Excel workbooks with PNG images .NET | validate image retention in combined workbook | C# Aspose.Cells picture handling during combine

using System;
using System.IO;
using Aspose.Cells;

// The example creates two workbooks, embeds a PNG image in each, merges the second workbook into the first using Workbook.Combine, iterates through all worksheets to count pictures, confirms that both images are retained, and saves the merged workbook as MergedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Paths to the images (replace with actual paths if needed)
            string imagePath1 = "image1.png";
            string imagePath2 = "image2.png";

            // Verify that image files exist before proceeding
            if (!File.Exists(imagePath1) || !File.Exists(imagePath2))
            {
                Console.WriteLine("One or both image files were not found. Please ensure the paths are correct.");
                return;
            }

            // Create the first workbook and embed the first image
            Workbook wb1 = new Workbook();
            // Add picture to cell A1 (row 0, column 0) using the file path overload
            wb1.Worksheets[0].Pictures.Add(0, 0, imagePath1);

            // Create the second workbook and embed the second image
            Workbook wb2 = new Workbook();
            wb2.Worksheets[0].Pictures.Add(0, 0, imagePath2);

            // Combine the second workbook into the first workbook
            wb1.Combine(wb2);

            // Count total pictures after combine
            int totalPictures = 0;
            foreach (Worksheet ws in wb1.Worksheets)
            {
                totalPictures += ws.Pictures.Count;
            }

            Console.WriteLine("Total pictures after combine: " + totalPictures);
            if (totalPictures == 2)
                Console.WriteLine("Embedded images are retained.");
            else
                Console.WriteLine("Embedded images are missing.");

            // Save the merged workbook
            string outputPath = "MergedWorkbook.xlsx";
            wb1.Save(outputPath);
            Console.WriteLine($"Merged workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
