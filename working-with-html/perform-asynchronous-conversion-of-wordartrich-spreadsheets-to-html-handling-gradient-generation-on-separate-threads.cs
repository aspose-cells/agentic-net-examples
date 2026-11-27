// Title: Asynchronously convert an Excel workbook with WordArt to HTML while applying two‑color gradients using Aspose.Cells for .NET
// Common Searches: c# async convert excel workbook containing WordArt to html with aspose.cells | apply two color gradient to WordArt shapes before html export using Aspose.Cells | parallel processing of shape gradients during Aspose.Cells HTML save | aspose.cells HtmlSaveOptions preserve WordArt and generate gradient images | run gradient generation on separate threads with Aspose.Cells in C#
// Tags: asynchronous excel to html conversion Aspose.Cells | WordArt gradient generation Aspose.Cells | two‑color gradient fill Shape Aspose.Cells | parallel task processing shape rendering Aspose.Cells | HtmlSaveOptions presentation preference Aspose.Cells | export workbook with WordArt to html Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WordArtHtmlAsyncConversion
{
    // The example loads an .xlsx workbook, finds all WordArt shapes, applies a two‑color horizontal gradient to each shape on background threads, waits for all gradient tasks to finish, and then saves the workbook as HTML using HtmlSaveOptions with presentation‑oriented settings.
    class Program
    {
        static async Task Main(string[] args)
        {
            // Validate Aspose.Cells license if you have one.
            // License license = new License();
            // license.SetLicense("Aspose.Cells.lic");

            string inputPath = args.Length > 0 ? args[0] : "WordArtSample.xlsx";
            string outputFolder = args.Length > 1 ? args[1] : "HtmlOutput";

            try
            {
                await ConvertWordArtWorkbookToHtmlAsync(inputPath, outputFolder);
                Console.WriteLine("Conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during conversion: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="excelPath">Path to the source .xlsx file.</param>
        /// <param name="htmlOutputFolder">Folder where HTML files and resources will be written.</param>
        private static async Task ConvertWordArtWorkbookToHtmlAsync(string excelPath, string htmlOutputFolder)
        {
            if (!File.Exists(excelPath))
                throw new FileNotFoundException($"Input file not found: {excelPath}");

            // Ensure output folder exists
            Directory.CreateDirectory(htmlOutputFolder);

            Workbook workbook;
            try
            {
                // Load workbook
                workbook = new Workbook(excelPath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to load workbook: {ex.Message}", ex);
            }

            // Collect tasks for gradient generation
            List<Task> gradientTasks = new List<Task>();

            // Iterate through all worksheets and their shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];
                    // Process only WordArt shapes
                    if (shape.IsWordArt)
                    {
                        Shape wordArtShape = shape; // capture for lambda
                        gradientTasks.Add(Task.Run(() => ApplyTwoColorGradient(wordArtShape)));
                    }
                }
            }

            // Await all gradient generation tasks
            await Task.WhenAll(gradientTasks);

            // Prepare HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Enable a more presentation‑oriented output (better CSS, inline styles, etc.)
                PresentationPreference = true,

                // Export each worksheet to a separate HTML file
                ExportActiveWorksheetOnly = false,

                // Keep images (including generated gradients) in a sub‑folder
                ExportImagesAsBase64 = false
                // ImageFolder property is not available in the current Aspose.Cells version;
                // the default image folder will be used.
            };

            // Save the workbook as HTML
            string htmlFilePath = Path.Combine(htmlOutputFolder,
                Path.GetFileNameWithoutExtension(excelPath) + ".html");

            try
            {
                workbook.Save(htmlFilePath, htmlOptions);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to save HTML: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Applies a two‑color gradient fill to the provided WordArt shape.
        /// This method runs on a background thread.
        /// </summary>
        /// <param name="shape">The WordArt shape to modify.</param>
        private static void ApplyTwoColorGradient(Shape shape)
        {
            // Choose gradient colors based on a simple hash of the shape name to get variety
            string name = shape.Name ?? Guid.NewGuid().ToString();
            int hash = name.GetHashCode();

            // Derive two colors from the hash
            Color startColor = Color.FromArgb(
                255,
                (hash & 0xFF0000) >> 16,
                (hash & 0x00FF00) >> 8,
                (hash & 0x0000FF));

            Color endColor = Color.FromArgb(
                255,
                (hash & 0x00FF00) >> 8,
                (hash & 0x0000FF),
                (hash & 0xFF0000) >> 16);

            // Apply the gradient using the FillFormat API
            shape.Fill.SetTwoColorGradient(startColor, endColor, GradientStyleType.Horizontal, 1);
        }
    }
}
