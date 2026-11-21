// Title: Export a single‑worksheet Excel workbook to HTML with separate CSS files using Aspose.Cells for .NET
// Common Searches: how to save an Excel workbook as HTML with external CSS files using Aspose.Cells C# | Aspose.Cells HtmlSaveOptions ExportWorksheetCSSSeparately example in .NET | C# generate HTML from a single sheet workbook and create a CSS folder with Aspose.Cells | disable base64 image embedding when exporting Excel to HTML with Aspose.Cells | ensure output directory exists before saving HTML using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions separate CSS files | C# Excel to HTML conversion external stylesheet | Aspose.Cells ExportImagesAsBase64 false | HTML export output directory handling Aspose.Cells | single worksheet HTML export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Creates a one‑sheet workbook, fills it with sample data, and saves it as HTML using HtmlSaveOptions with ExportWorksheetCSSSeparately enabled and ExportImagesAsBase64 disabled, while ensuring the target directory exists for the generated CSS folder.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook with a single worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate the worksheet with sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apples");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Bananas");
                sheet.Cells["B3"].PutValue(85);
                sheet.Cells["A4"].PutValue("Cherries");
                sheet.Cells["B4"].PutValue(60);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export each worksheet's CSS to a separate file
                    ExportWorksheetCSSSeparately = true,

                    // Optional: do not embed images as Base64 (they will be saved as separate files)
                    ExportImagesAsBase64 = false
                };

                // Define output path
                string htmlFilePath = "Workbook.html";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(htmlFilePath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML with the specified options
                workbook.Save(htmlFilePath, htmlOptions);

                Console.WriteLine("Workbook exported to HTML with separate CSS files.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
