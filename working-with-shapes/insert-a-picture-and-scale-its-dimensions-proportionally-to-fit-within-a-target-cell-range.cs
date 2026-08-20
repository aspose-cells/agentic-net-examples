// Title: C# – Insert and proportionally scale a picture to fit a cell range with Aspose.Cells
// Description: Creates a new workbook, defines a target range (e.g., B2:E10), loads an image, adds it to the worksheet using the Add method with cell coordinates, sets the picture to be placed inside the cells, locks its aspect ratio, and saves the file as an XLSX document.
// Keywords: Aspose.Cells picture placement | C# insert image into cells | scale image to fit cell range Aspose.Cells | lock aspect ratio picture Aspose.Cells | add picture with cell boundaries .NET | FitImageToCellRange Aspose.Cells
// Common Searches: Aspose.Cells add picture to specific cell range C# | fit image proportionally inside cells Aspose.Cells | place picture in cells and lock aspect ratio .NET | resize picture to match cell dimensions Aspose.Cells | C# Aspose.Cells picture scaling example
// Developer Intent: Insert an image and have it automatically resize to fill a defined cell range while preserving its original aspect ratio.
// Use Cases: Add a company logo across cells B2:E10 in a generated report without distortion. | Insert product photos into catalog cells, ensuring each image scales proportionally. | Automate placement of scanned signatures in a form area, keeping signature proportions intact.
// AI Prompts: Show C# code to insert a picture into a worksheet range and maintain aspect ratio using Aspose.Cells. | Provide an Aspose.Cells example that fits an image inside cells B2:E10 and locks its aspect ratio. | Explain how to adjust picture placement and scaling when the target cell range changes in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureFitExample
{
    // Creates a new workbook, defines a target range (e.g., B2:E10), loads an image, adds it to the worksheet using the Add method with cell coordinates, sets the picture to be placed inside the cells, locks its aspect ratio, and saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the target cell range where the picture should fit (e.g., B2:E10)
                int topRow = 1;      // B2 -> row index 1 (zero‑based)
                int leftColumn = 1;  // B2 -> column index 1
                int bottomRow = 9;   // E10 -> row index 9
                int rightColumn = 4; // E10 -> column index 4

                string imagePath = "sample.jpg";

                // Ensure the image file exists before attempting to load it
                if (File.Exists(imagePath))
                {
                    using (FileStream imageStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                    {
                        // Add the picture to the worksheet within the specified cell range.
                        int pictureIndex = worksheet.Pictures.Add(topRow, leftColumn, bottomRow, rightColumn, imageStream);

                        // Retrieve the picture object
                        Picture picture = worksheet.Pictures[pictureIndex];

                        // Place the picture inside the cells and lock aspect ratio
                        picture.IsPlacedInCell = true;
                        picture.IsAspectRatioLocked = true; // Obsolete but still functional
                    }
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Save the workbook
                string outputPath = "PictureFitInRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
