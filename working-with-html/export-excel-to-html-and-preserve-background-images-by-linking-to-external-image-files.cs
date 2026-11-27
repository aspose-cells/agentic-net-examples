// Title: Export an Excel workbook to HTML with linked background images using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to HTML without embedding images | How to keep worksheet background image as separate file when saving as HTML with Aspose.Cells | Save Excel as HTML with external image files using HtmlSaveOptions in .NET | Export Excel to HTML and generate _files folder for images with Aspose.Cells | C# Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 false example
// Tags: HtmlSaveOptions external image linking | Aspose.Cells preserve worksheet background | Excel to HTML image folder generation | C# export images as separate files Aspose.Cells | HTML export without base64 encoding Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace ExcelToHtmlWithBackgroundImages
{
    // The sample loads an Excel file with Aspose.Cells, sets HtmlSaveOptions to disable base64 image embedding, and saves the workbook as HTML. All images, including any worksheet background, are written to a separate "_files" folder, producing HTML that references external image files.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust them if you place the files elsewhere.
            string inputExcelPath = "input.xlsx";
            string outputHtmlPath = "output.html";

            // Validate input file existence.
            if (!File.Exists(inputExcelPath))
            {
                Console.WriteLine($"Error: Input file \"{inputExcelPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputExcelPath);

                // OPTIONAL: Add a background image if needed.
                // The BackgroundImage class may not be available in some versions,
                // so this block is omitted to ensure compilation.
                // If your version supports it, you can uncomment and adjust accordingly.

                // Configure HTML save options.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Do NOT embed images as base64; we want separate image files.
                    ExportImagesAsBase64 = false,

                    // Keep the original worksheet layout (gridlines, etc.) – adjust as needed.
                    ExportGridLines = true
                };

                // Save the workbook as HTML. Images (including the background) will be written
                // to a folder named "<outputHtmlPath>_files" next to the HTML file.
                workbook.Save(outputHtmlPath, htmlOptions);

                Console.WriteLine("HTML export completed successfully.");
                Console.WriteLine($"HTML file: {Path.GetFullPath(outputHtmlPath)}");
                Console.WriteLine($"Images folder: {Path.GetFullPath(outputHtmlPath + "_files")}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the export process:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
