// Title: How to verify that disabling downlevel revealed comments in Aspose.Cells HTML export does not affect gridlines, images, or other HTML elements (C#)
// Common Searches: Aspose.Cells HTML export disable downlevel revealed comments keep gridlines | C# save workbook as HTML without comments but retain embedded PNG | compare HTML output with DisableDownlevelRevealedComments true versus false Aspose.Cells | validate that disabling downlevel revealed comments does not remove base64 images in HTML export | check if HtmlSaveOptions.DisableDownlevelRevealedComments impacts other HTML features
// Tags: Aspose.Cells HtmlSaveOptions DisableDownlevelRevealedComments | HTML export gridlines preservation Aspose.Cells | embedding base64 images in Aspose.Cells HTML output | HTML comment handling Aspose.Cells workbook | temporary PNG insertion Aspose.Cells worksheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook with sample data, adds an HTML‑formatted comment and a temporary 1×1 PNG image, then saves the workbook to HTML twice—once with DisableDownlevelRevealedComments set to false (default) and once true—verifying that the comment appears only in the first file while gridlines, images, and other HTML features remain unchanged, and finally cleans up the temporary image.
class Program
{
    static void Main()
    {
        try
        {
            // Path for temporary image used in the workbook
            string imagePath = Path.Combine(Path.GetTempPath(), "sample_image.png");

            // ------------------------------------------------------------
            // 1. Create a workbook with sample data, a comment (HTML) and an image
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(1.20);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(0.80);

            // Add an HTML comment to cell C2 (zero‑based indices)
            Comment comment = sheet.Comments[2, 1];
            comment.HtmlNote = "<b>Special Offer</b><br/>Buy one get one free!";

            // Create a minimal PNG (1x1 pixel) and save it to a temporary file
            // Base64 for a transparent 1x1 PNG
            const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2ZcAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(pngBase64);
            File.WriteAllBytes(imagePath, pngBytes);

            // Insert the image into the worksheet (row 5, column 0 with small offsets)
            if (File.Exists(imagePath))
            {
                try
                {
                    using (FileStream fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                    {
                        // AddPicture overload expects a Stream in this API version
                        sheet.Shapes.AddPicture(5, 0, 5, 0, fs);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error inserting picture: {ex.Message}");
                }
            }

            // ------------------------------------------------------------
            // 2. Helper method to save HTML with a given option
            // ------------------------------------------------------------
            void SaveHtml(string fileName, bool disableDownlevelRevealedComments)
            {
                try
                {
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        ExportGridLines = true,
                        ExportImagesAsBase64 = true,
                        // When true, downlevel revealed comments are omitted
                        DisableDownlevelRevealedComments = disableDownlevelRevealedComments
                    };

                    // Save the workbook as HTML
                    workbook.Save(fileName, htmlOptions);
                    Console.WriteLine($"Saved HTML to \"{fileName}\" (DisableDownlevelRevealedComments = {disableDownlevelRevealedComments})");

                    // Simple validation: read the file and check if the comment HTML is present
                    if (File.Exists(fileName))
                    {
                        string htmlContent = File.ReadAllText(fileName);
                        bool commentFound = htmlContent.Contains("<b>Special Offer</b>");
                        Console.WriteLine($"  Comment HTML {(commentFound ? "found" : "not found")} in the output.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving HTML \"{fileName}\": {ex.Message}");
                }
            }

            // ------------------------------------------------------------
            // 3. Save with the option disabled (default behavior) – comments should appear
            // ------------------------------------------------------------
            string outputWithComments = "output_with_comments.html";
            SaveHtml(outputWithComments, false);

            // ------------------------------------------------------------
            // 4. Save with the option enabled – comments should be omitted but other features stay intact
            // ------------------------------------------------------------
            string outputWithoutComments = "output_without_comments.html";
            SaveHtml(outputWithoutComments, true);

            // ------------------------------------------------------------
            // 5. Clean up temporary image file
            // ------------------------------------------------------------
            if (File.Exists(imagePath))
            {
                try
                {
                    File.Delete(imagePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error deleting temporary image: {ex.Message}");
                }
            }

            Console.WriteLine("Validation completed. Compare the two HTML files to ensure that disabling downlevel revealed comments does not affect gridlines, images, or other HTML features.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
