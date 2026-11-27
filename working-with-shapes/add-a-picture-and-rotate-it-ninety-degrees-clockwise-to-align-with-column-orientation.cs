// Title: Insert an image into a specific cell and rotate it 90° clockwise with Aspose.Cells for .NET
// Common Searches: C# code to add a picture to cell B2 in an Excel workbook using Aspose.Cells | how to rotate an inserted picture 90 degrees clockwise in Aspose.Cells .NET | Aspose.Cells example for setting picture rotation angle programmatically | adjust picture size to match row height and span multiple columns with Aspose.Cells C# | command‑line tool to embed and rotate an image in Excel using Aspose.Cells
// Tags: insert picture into cell Aspose.Cells | picture rotation angle Aspose.Cells | match picture height to row Aspose.Cells | stretch picture across columns Aspose.Cells | save workbook after picture changes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AddAndRotatePicture
{
    // The sample validates command‑line arguments, creates a new Workbook, inserts an image into cell B2 of the first worksheet, sets the picture's RotationAngle to 90°, optionally aligns its height to the row and expands its width across three columns, and saves the workbook to the specified output file.
    class Program
    {
        /// <param name="args">
        /// args[0] - Path to the source image file (required).
        /// args[1] - Path to the output Excel file (optional, defaults to "RotatedPicture.xlsx").
        /// </param>
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args == null || args.Length == 0 || string.IsNullOrWhiteSpace(args[0]))
            {
                Console.WriteLine("Usage: AddAndRotatePicture <ImagePath> [OutputExcelPath]");
                return;
            }

            string imagePath = args[0];
            string outputPath = args.Length > 1 && !string.IsNullOrWhiteSpace(args[1])
                ? args[1]
                : "RotatedPicture.xlsx";

            // Verify that the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Error: Image file not found at \"{imagePath}\"");
                return;
            }

            // Ensure the output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error preparing output directory: {ex.Message}");
                return;
            }

            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // 3. Add the picture to cell B2 (row 1, column 1)
                int pictureRow = 1;   // zero‑based index (B2)
                int pictureColumn = 1;
                int pictureIndex = sheet.Pictures.Add(pictureRow, pictureColumn, imagePath);
                Picture picture = sheet.Pictures[pictureIndex];

                // 4. Rotate the picture 90 degrees clockwise
                picture.RotationAngle = 90.0;

                // 5. Adjust picture size (optional)
                // Set height to match the row height
                picture.Height = (int)sheet.Cells.Rows[pictureRow].Height;
                // Set width to span three columns
                picture.Width = (int)(sheet.Cells.Columns[pictureColumn].Width * 3);

                // 6. Save the workbook
                workbook.Save(outputPath);

                Console.WriteLine($"Picture added and rotated successfully. Saved to \"{outputPath}\"");
            }
            catch (Exception ex)
            {
                // Report full exception details
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
