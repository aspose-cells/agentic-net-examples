// Title: Copy All Pictures from Multiple Worksheets to a Summary Sheet (Preserve Size) – Aspose.Cells for .NET C# Example
// Description: Demonstrates how to create a workbook, add images to several worksheets, iterate through each sheet (excluding the summary), retrieve every picture with GetAllPictures, copy the image data to a new summary sheet using the original anchor coordinates, and preserve the original dimensions and formatting with Picture.Copy and CopyOptions before saving the file.
// Keywords: Aspose.Cells copy pictures C# | preserve image size Excel | batch copy images Aspose.Cells | GetAllPictures example | summary worksheet picture aggregation | Excel image duplication .NET | copy shapes between sheets
// Common Searches: copy all images from worksheets to one sheet Aspose.Cells | preserve picture dimensions when moving Excel images | Aspose.Cells C# example to aggregate pictures | how to duplicate pictures across worksheets in .NET | batch copy pictures in Excel using Aspose
// Developer Intent: Transfer every picture from each worksheet into a single summary sheet while keeping its original size and formatting.
// Use Cases: Create a catalog sheet that gathers product photos from multiple category tabs. | Compile design mockups from several project worksheets into one overview page. | Migrate embedded images from legacy Excel files into a new template that displays all images together.
// AI Prompts: Generate C# code with Aspose.Cells that loops through all worksheets, copies each picture to a summary sheet, and retains original size and formatting. | Show how to extract Picture.Data, add the image to another worksheet using the same anchor coordinates, and use CopyOptions to preserve properties. | Explain why Picture.Copy with default CopyOptions is needed to keep image dimensions when duplicating pictures across sheets.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add images to several worksheets, iterate through each sheet (excluding the summary), retrieve every picture with GetAllPictures, copy the image data to a new summary sheet using the original anchor coordinates, and preserve the original dimensions and formatting with Picture.Copy and CopyOptions before saving the file.
    public class BatchCopyPicturesToSummary
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Sample data: add two worksheets with pictures
                // -------------------------------------------------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";

                // Add a picture to Sheet1 if the file exists
                string imgPath1 = "image1.jpg";
                if (File.Exists(imgPath1))
                {
                    sheet1.Pictures.Add(2, 2, imgPath1);
                }

                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                // Add a picture to Sheet2 if the file exists
                string imgPath2 = "image2.png";
                if (File.Exists(imgPath2))
                {
                    sheet2.Pictures.Add(5, 3, imgPath2);
                }

                // -------------------------------------------------
                // Create (or get) the summary worksheet where all pictures will be copied
                // -------------------------------------------------
                Worksheet summarySheet = workbook.Worksheets.Add("Summary");

                // -------------------------------------------------
                // Iterate through all worksheets except the summary sheet
                // -------------------------------------------------
                for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
                {
                    Worksheet ws = workbook.Worksheets[wsIndex];
                    if (ws.Name == summarySheet.Name) continue; // skip summary sheet

                    // Get all pictures from the current worksheet
                    Picture[] pictures = ws.GetAllPictures(); // includes embedded and floating pictures

                    foreach (Picture srcPic in pictures)
                    {
                        // Retrieve picture binary data
                        byte[] imgData = srcPic.Data;
                        if (imgData == null || imgData.Length == 0) continue; // safety check

                        // Preserve original position and size using the picture's anchor coordinates
                        int topRow = srcPic.UpperLeftRow;
                        int leftColumn = srcPic.UpperLeftColumn;
                        int bottomRow = srcPic.LowerRightRow;
                        int rightColumn = srcPic.LowerRightColumn;

                        // Add a placeholder picture to the summary sheet using the same image data
                        using (MemoryStream ms = new MemoryStream(imgData))
                        {
                            int newPicIdx = summarySheet.Pictures.Add(topRow, leftColumn, bottomRow, rightColumn, ms);
                            Picture destPic = summarySheet.Pictures[newPicIdx];

                            // Copy all properties from source picture to destination picture
                            CopyOptions copyOptions = new CopyOptions(); // default options
                            destPic.Copy(srcPic, copyOptions); // preserves original size and formatting
                        }
                    }
                }

                // -------------------------------------------------
                // Save the workbook (lifecycle: save)
                // -------------------------------------------------
                string outPath = "BatchCopyPicturesSummary.xlsx";
                workbook.Save(outPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            BatchCopyPicturesToSummary.Run();
        }
    }
}
