// Title: Aspose.Cells C# – Export Workbook to HTML with Image Alt Text Matching Cell Address
// Description: This example verifies a JPEG file, creates a workbook, writes headers to A1 and B1, inserts a picture anchored at B2, assigns the picture’s AlternativeText to the cell name (e.g., "B2"), configures HtmlSaveOptions to embed images as Base64 and to add an id attribute containing the cell reference to each <td>, and saves the result as a single HTML document.
// Keywords: Aspose.Cells HTML export C# | image alt text cell reference | ExportImagesAsBase64 | CellNameAttribute id | accessible Excel to HTML conversion | embed pictures base64 Aspose | C# workbook to HTML with accessibility
// Common Searches: set picture alt attribute to cell address Aspose.Cells | HTML export with base64 images from Excel C# | add id to table cells based on Excel cell name | accessible HTML output from Aspose.Cells workbook | how to anchor image to specific cell in HTML export
// Developer Intent: Generate an HTML file from a workbook where every embedded picture carries an alt attribute equal to its originating cell reference, ensuring accessibility and easy DOM targeting.
// Use Cases: Create self‑contained HTML reports with product photos that screen readers can identify by cell location. | Enable CSS or JavaScript to target individual cells using id attributes derived from Excel cell names. | Distribute a single HTML file without external image resources while preserving accessibility metadata.
// AI Prompts: Show how to prepend custom text to the alt attribute, e.g., "Product image at B2". | Demonstrate adding a title attribute to each <td> element via HtmlSaveOptions. | Explain how to export images as separate files instead of Base64 while still setting alt text to the cell address.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example verifies a JPEG file, creates a workbook, writes headers to A1 and B1, inserts a picture anchored at B2, assigns the picture’s AlternativeText to the cell name (e.g., "B2"), configures HtmlSaveOptions to embed images as Base64 and to add an id attribute containing the cell reference to each <td>, and saves the result as a single HTML document.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the image file exists to avoid FileNotFoundException
            const string imagePath = "sample.jpg";
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file \"{imagePath}\" not found.");
                return;
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Image");

            // Add an image anchored to cell B2 (row index 1, column index 1)
            int imageRow = 1; // zero‑based index for row 2
            int imageCol = 1; // zero‑based index for column B

            // Add picture and retrieve the Picture object
            int pictureIndex = worksheet.Pictures.Add(imageRow, imageCol, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Set the alt text of the image to the cell address (e.g., "B2") for accessibility
            picture.AlternativeText = worksheet.Cells[imageRow, imageCol].Name; // "B2"

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Embed images directly in the HTML as Base64 strings
                ExportImagesAsBase64 = true,
                // Add a cell identifier attribute (e.g., id="B2") to each <td>
                CellNameAttribute = "id"
            };

            // Save the workbook as an HTML file
            const string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
