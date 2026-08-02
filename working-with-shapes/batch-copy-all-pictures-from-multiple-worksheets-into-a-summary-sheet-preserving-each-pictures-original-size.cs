// Title: Copy All Pictures from Multiple Worksheets to a Summary Sheet (Preserve Size) – Aspose.Cells for .NET
// Description: This Aspose.Cells for .NET example creates a workbook, adds images to several worksheets, then generates a "Summary" sheet. It iterates through every worksheet, extracts each picture's raw data via a memory stream, and inserts the image onto the summary sheet while keeping the original height, width, border color, border weight, and placement. Pictures are spaced by a configurable row offset before the workbook is saved.
// Keywords: Aspose.Cells copy pictures | C# copy Excel images | preserve picture dimensions | batch copy images Excel | summary worksheet pictures | .NET Excel shape handling | Aspose.Cells picture properties | Excel image aggregation
// Common Searches: how to copy all images from Excel worksheets using Aspose.Cells | preserve picture size when moving pictures between sheets .NET | batch transfer pictures to a summary sheet Aspose.Cells | C# example copy Excel pictures to another worksheet | Aspose.Cells copy picture properties
// Developer Intent: Transfer every picture from each worksheet into a single summary sheet while retaining the original size and formatting.
// Use Cases: Create a catalog sheet that aggregates product photos from department worksheets. | Build a visual dashboard that gathers thumbnails from multiple project sheets for quick review. | Migrate embedded graphics from legacy worksheets into a centralized summary page in a reporting workbook.
// AI Prompts: Generate C# code with Aspose.Cells that copies all pictures from every worksheet to a new summary sheet, preserving dimensions and borders. | Show how to iterate through worksheets, extract picture data, and add each image to a summary sheet with adjustable row spacing using Aspose.Cells. | Explain safe handling of pictures that may have missing or empty image data when copying between worksheets in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This Aspose.Cells for .NET example creates a workbook, adds images to several worksheets, then generates a "Summary" sheet. It iterates through every worksheet, extracts each picture's raw data via a memory stream, and inserts the image onto the summary sheet while keeping the original height, width, border color, border weight, and placement. Pictures are spaced by a configurable row offset before the workbook is saved.
    public class BatchCopyPicturesToSummary
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add first worksheet and a picture (ensure the image file exists)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            string imgPath1 = "image1.jpg";
            if (File.Exists(imgPath1))
                sheet1.Pictures.Add(2, 2, imgPath1);
            else
                Console.WriteLine($"Warning: '{imgPath1}' not found. Skipping picture addition.");

            // Add second worksheet and a picture
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            string imgPath2 = "image2.png";
            if (File.Exists(imgPath2))
                sheet2.Pictures.Add(4, 4, imgPath2);
            else
                Console.WriteLine($"Warning: '{imgPath2}' not found. Skipping picture addition.");

            // Add a summary worksheet where all pictures will be copied
            Worksheet summarySheet = workbook.Worksheets.Add("Summary");

            // Positioning variables for copied pictures
            int currentRow = 0;
            const int rowSpacing = 20; // rows to skip between pictures

            // Iterate through all worksheets except the summary sheet
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name == summarySheet.Name)
                    continue;

                // Iterate through each picture in the current worksheet
                foreach (Picture srcPic in ws.Pictures)
                {
                    // Retrieve raw image data
                    byte[] imgData = srcPic.Data;
                    if (imgData == null || imgData.Length == 0)
                        continue; // safety check

                    // Add picture to the summary sheet using a memory stream
                    using (MemoryStream ms = new MemoryStream(imgData))
                    {
                        int picIndex = summarySheet.Pictures.Add(currentRow, 0, ms);
                        Picture destPic = summarySheet.Pictures[picIndex];

                        // Preserve original size
                        destPic.Height = srcPic.Height;
                        destPic.Width = srcPic.Width;

                        // Preserve visual properties
                        destPic.BorderLineColor = srcPic.BorderLineColor;
                        destPic.BorderWeight = srcPic.BorderWeight;
                        destPic.Placement = srcPic.Placement;

                        // Move to next row position for the following picture
                        currentRow += rowSpacing;
                    }
                }
            }

            // Save the workbook
            string outputPath = "BatchCopyPicturesSummary.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
