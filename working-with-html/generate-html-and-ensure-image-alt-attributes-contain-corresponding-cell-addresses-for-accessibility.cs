// Title: C# – Export Excel to Accessible HTML with Cell IDs and Image Alt Text using Aspose.Cells
// Description: This example creates a workbook, populates product rows, inserts a picture into cell A5 (when the file exists), assigns the cell address as the image's alt attribute, and saves the sheet as a self‑contained HTML file. The HtmlSaveOptions are set to output cell coordinates, use the address as the element ID, and embed images as Base64 strings, delivering an accessible, portable web page.
// Keywords: Aspose.Cells C# HTML export | Excel to HTML with cell IDs | image alt attribute Aspose.Cells | ExportCellCoordinate option | CellNameAttribute id | ExportImagesAsBase64 | accessible HTML from Excel | screen‑reader friendly markup | self‑contained HTML report | sample code .NET
// Common Searches: how to add alt text to images when exporting Excel to HTML with Aspose.Cells | export cell address as HTML element id using Aspose.Cells .NET | embed pictures as base64 in HTML output from Aspose.Cells | Aspose.Cells HtmlSaveOptions for accessibility | C# code to generate HTML with data‑celladdress attribute
// Developer Intent: Produce an HTML document from a workbook where each table cell carries a unique identifier and any embedded picture includes an alt attribute that matches its originating cell.
// Use Cases: Generate web‑ready reports that comply with WCAG guidelines by providing identifiable cells and descriptive alt text. | Create email‑friendly HTML snippets that contain all images inline, eliminating external dependencies. | Build interactive dashboards where JavaScript can target specific cells via their IDs for dynamic updates.
// AI Prompts: Show how to set different alt texts for multiple pictures based on their cell locations. | Describe the HTML markup changes introduced by ExportCellCoordinate and CellNameAttribute settings. | Give a code sample that adds a custom data‑celladdress attribute while preserving existing cell formatting.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsHtmlExport
{
    // This example creates a workbook, populates product rows, inserts a picture into cell A5 (when the file exists), assigns the cell address as the image's alt attribute, and saves the sheet as a self‑contained HTML file. The HtmlSaveOptions are set to output cell coordinates, use the address as the element ID, and embed images as Base64 strings, delivering an accessible, portable web page.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add some sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.20);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(0.80);

                // Insert an image into cell A5 if the file exists
                const string imagePath = "sample-image.png";
                if (File.Exists(imagePath))
                {
                    // Row 4 (zero‑based) corresponds to A5, column 0 is column A
                    int pictureIndex = sheet.Pictures.Add(4, 0, imagePath);
                    Picture pic = sheet.Pictures[pictureIndex];
                    // The AlternativeText property becomes the alt attribute in the generated HTML
                    pic.AlternativeText = "A5";
                }
                else
                {
                    Console.WriteLine($"Image file \"{imagePath}\" not found. Skipping picture insertion.");
                }

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export cell coordinates as an attribute (e.g., data-celladdress) for accessibility
                    ExportCellCoordinate = true,
                    // Use the cell address as the HTML element id (e.g., <td id=\"A5\">)
                    CellNameAttribute = "id",
                    // Embed images directly as Base64 strings so the <img> tag appears in the HTML
                    ExportImagesAsBase64 = true
                };

                // Save the workbook as an HTML file
                const string outputPath = "output.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"HTML file \"{outputPath}\" generated with cell coordinates and image alt text.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
