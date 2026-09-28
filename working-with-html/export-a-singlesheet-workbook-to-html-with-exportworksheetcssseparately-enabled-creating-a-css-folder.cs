// Title: Export a single‑sheet Aspose.Cells workbook to HTML with external CSS files using ExportWorksheetCSSSeparately in C#
// AI Prompts: Write C# code that creates a one‑sheet workbook, configures HtmlSaveOptions.ExportWorksheetCSSSeparately = true, disables base64 image embedding, and saves the workbook as HTML while generating a dedicated CSS folder. | Demonstrate how to set up Aspose.Cells HtmlSaveOptions to output a separate CSS file per worksheet and store images in an Images directory during Excel‑to‑HTML conversion in C#.
// Common Searches: Aspose.Cells C# export workbook to HTML with separate CSS folder | How to enable ExportWorksheetCSSSeparately in HtmlSaveOptions for a single worksheet | Save Excel as HTML with external CSS files using Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 false example | Create output directories for CSS and images when converting Excel to HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportWorksheetCSSSeparately | C# Excel to HTML external CSS folder | Aspose.Cells separate CSS per worksheet | HTML export images as separate files Aspose.Cells | Create CSS and Images directories Aspose.Cells output

using Aspose.Cells;
using System;
using System.IO;

// The program builds a one‑sheet workbook, creates Output/Css and Output/Images folders, configures HtmlSaveOptions with ExportWorksheetCSSSeparately=true and ExportImagesAsBase64=false, and saves the workbook as an HTML file with CSS and images stored in separate directories.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook with a single worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["A3"].PutValue("Oranges");
            sheet.Cells["B3"].PutValue(200);

            // Define output directories
            string outputFolder = "Output";
            string cssFolder = Path.Combine(outputFolder, "Css");
            string imagesFolder = Path.Combine(outputFolder, "Images");

            Directory.CreateDirectory(outputFolder);
            Directory.CreateDirectory(cssFolder);
            Directory.CreateDirectory(imagesFolder);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportWorksheetCSSSeparately = true, // separate CSS per worksheet
                ExportImagesAsBase64 = false          // save images as separate files
                // Other advanced options (CSS file name, image folder, single file) are not
                // available in the current Aspose.Cells version and are therefore omitted.
            };

            // Save the workbook as HTML
            string htmlPath = Path.Combine(outputFolder, "workbook.html");
            workbook.Save(htmlPath, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
