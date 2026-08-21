// Title: Enable CSS Custom Properties & TableCssId in Aspose.Cells HtmlSaveOptions (C#)
// Description: Demonstrates how to export an Excel workbook to HTML with Aspose.Cells for .NET while activating CSS custom properties and assigning a TableCssId. The example creates a sample sheet, applies header styling, optionally embeds a logo image, defines inline CSS rules linked to the TableCssId, and saves the file as HTML with enhanced table‑style flexibility.
// Keywords: Aspose.Cells | HtmlSaveOptions | EnableCssCustomProperties | TableCssId | C# | custom CSS | HTML export | Excel to HTML | CSS variables | table styling | workbook.Save | .NET
// Common Searches: Aspose.Cells enable CSS custom properties when saving to HTML | How to set TableCssId in HtmlSaveOptions | Export Excel to HTML with custom table CSS using Aspose.Cells | C# HtmlSaveOptions CssStyles example | Aspose.Cells HTML export custom styling
// Developer Intent: Activate CSS custom properties and assign a TableCssId in HtmlSaveOptions to apply user‑defined CSS rules to the generated HTML table.
// Use Cases: Create a reusable visual theme for all exported tables by linking them to a specific TableCssId and providing matching CSS rules. | Leverage CSS custom properties for dynamic theming (colors, spacing, borders) in the HTML output. | Combine worksheet image insertion with styled HTML export while preserving custom table formatting.
// AI Prompts: Generate a CssStyles block that uses CSS variables for table background and border colors with EnableCssCustomProperties enabled. | Show how to export multiple worksheets to separate HTML files, each with a different TableCssId and corresponding CSS. | Explain how to override the CSS variables defined by EnableCssCustomProperties in an external stylesheet after the HTML file is saved.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to export an Excel workbook to HTML with Aspose.Cells for .NET while activating CSS custom properties and assigning a TableCssId. The example creates a sample sheet, applies header styling, optionally embeds a logo image, defines inline CSS rules linked to the TableCssId, and saves the file as HTML with enhanced table‑style flexibility.
    public class HtmlSaveOptionsEnableCssCustomPropertiesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data with formatting to demonstrate CSS styling
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Price");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(1.25);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(0.80);

                // Apply style to the header row
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.Font.Color = Color.White;
                headerStyle.ForegroundColor = Color.DarkBlue;
                headerStyle.Pattern = BackgroundType.Solid;
                worksheet.Cells["A1"].SetStyle(headerStyle);
                worksheet.Cells["B1"].SetStyle(headerStyle);

                // Path to the image file
                const string imagePath = "logo.png";

                // Insert the same image into two different cells if the file exists
                if (File.Exists(imagePath))
                {
                    int pictureIndex1 = worksheet.Pictures.Add(2, 0, imagePath);
                    Picture picture1 = worksheet.Pictures[pictureIndex1];
                    picture1.Width = 50;
                    picture1.Height = 50;

                    int pictureIndex2 = worksheet.Pictures.Add(4, 0, imagePath);
                    Picture picture2 = worksheet.Pictures[pictureIndex2];
                    picture2.Width = 50;
                    picture2.Height = 50;
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    EnableCssCustomProperties = true,
                    TableCssId = "custom-table-style",
                    CssStyles = @"
                        #custom-table-style table { border: 1px solid #ccc; }
                        #custom-table-style th { background-color: #f2f2f2; }
                        #custom-table-style td { padding: 5px; }"
                };

                // Save the workbook as HTML using the configured options
                const string outputHtml = "HtmlWithCssCustomPropertiesAndTableCssId.html";
                workbook.Save(outputHtml, htmlOptions);

                Console.WriteLine($"HTML file saved as '{outputHtml}' with EnableCssCustomProperties=true and TableCssId set.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    internal class Program
    {
        private static void Main(string[] args)
        {
            HtmlSaveOptionsEnableCssCustomPropertiesDemo.Run();
        }
    }
}
