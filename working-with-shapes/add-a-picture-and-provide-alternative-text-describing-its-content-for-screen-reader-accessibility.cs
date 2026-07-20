// Title: C# – Insert a picture into an Excel worksheet and set alt text with Aspose.Cells
// Description: This example creates a new Workbook, verifies a PNG file, inserts the image into cells B2‑F5 using Worksheet.Pictures.Add, assigns descriptive AlternativeText for screen‑reader accessibility, and saves the file as an .xlsx workbook.
// Keywords: Aspose.Cells C# picture insert | Excel image alternative text | set AlternativeText Aspose.Cells | accessibility Excel Aspose | add image to worksheet .NET | Worksheet.Pictures.Add example | screen reader alt text Excel | Aspose.Cells sample code | C# Excel image insertion
// Common Searches: how to add a picture to Excel with Aspose.Cells C# | set alternative text for picture in Aspose.Cells | Aspose.Cells picture accessibility example | insert image into specific cell range Aspose.Cells | C# code to add PNG to worksheet and set alt text
// Developer Intent: Add an image to a worksheet and provide alt text for accessibility using Aspose.Cells.
// Use Cases: Include a company logo in generated reports with alt text so screen readers can describe it. | Embed chart screenshots in financial models and supply concise descriptions for compliance with accessibility standards. | Programmatically populate a product catalog worksheet with photos, each tagged with unique alternative text for visually impaired users.
// AI Prompts: Write C# code that loads an image from a URL, inserts it into an Excel sheet with Aspose.Cells, and sets AlternativeText for accessibility. | Show how to iterate over a folder of images, add each to a separate worksheet range, and assign distinct AlternativeText values. | Explain best practices for handling missing image files when adding pictures with Aspose.Cells in C#.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a new Workbook, verifies a PNG file, inserts the image into cells B2‑F5 using Worksheet.Pictures.Add, assigns descriptive AlternativeText for screen‑reader accessibility, and saves the file as an .xlsx workbook.
    class AddPictureWithAltText
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image file to be inserted
                string imagePath = "sample.png"; // replace with your actual image file path

                // Verify that the image file exists
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Add the picture to the worksheet (from row 1, column 1 to row 5, column 5)
                int pictureIndex = worksheet.Pictures.Add(1, 1, 5, 5, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Set alternative text for screen reader accessibility
                picture.AlternativeText = "Company logo showing a blue circle with white text";

                // Save the workbook
                string outputPath = "PictureWithAltText.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point
        static void Main(string[] args)
        {
            Run();
        }
    }
}
