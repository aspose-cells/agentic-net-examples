using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Verify that the image file exists before attempting to load it.
                const string imagePath = "sample.jpg";
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {Path.GetFullPath(imagePath)}");
                    return;
                }

                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the target cell range where the picture should fit (e.g., B2:E5).
                // Row and column indices are zero‑based.
                int topRow = 1;      // B2 -> row index 1
                int leftColumn = 1;  // B2 -> column index 1
                int bottomRow = 4;   // E5 -> row index 4
                int rightColumn = 4; // E5 -> column index 4

                // Load the image into a stream and add the picture to the worksheet.
                using (FileStream imageStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                {
                    // The picture will be scaled proportionally to fit within the defined range.
                    worksheet.Pictures.Add(topRow, leftColumn, bottomRow, rightColumn, imageStream);
                }

                // Save the workbook.
                const string outputPath = "PictureInRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}