// Title: Create a separate external CSS file for each worksheet when converting an Excel workbook to HTML with Aspose.Cells for .NET (C#)
// Common Searches: how to save each Excel worksheet as an HTML file with its own CSS using Aspose.Cells C# | Aspose.Cells C# generate per‑sheet external stylesheet during HTML export | C# convert workbook to HTML and create individual CSS files for each sheet | export active worksheet only to HTML with separate CSS file Aspose.Cells | Aspose.Cells HtmlSaveOptions external CSS per worksheet example C#
// Tags: Aspose.Cells HtmlSaveOptions per‑worksheet external CSS | C# export Excel worksheet to HTML with separate stylesheet | Aspose.Cells generate individual CSS files for each sheet | HTML conversion of workbook with distinct CSS per sheet | C# save workbook as HTML using external CSS files

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel file, creates an output folder, and iterates through each worksheet. For every sheet it sets HtmlSaveOptions to export only the active worksheet and to keep images as separate files, then saves the sheet as HTML. Aspose.Cells automatically creates a dedicated external CSS file (and image folder) for each worksheet, providing per‑sheet styling while handling errors individually.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the existing Excel file.
            Workbook workbook = new Workbook(inputPath);

            // Directory where HTML, CSS and image files will be saved.
            string outputDir = "output";
            Directory.CreateDirectory(outputDir);

            // Process each worksheet individually.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Configure HTML save options for external CSS and separate images.
                    HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
                    {
                        ExportActiveWorksheetOnly = true,   // Export only the current sheet.
                        ExportImagesAsBase64 = false        // Store images as separate files.
                        // ImageFolder is omitted; Aspose will create a default folder for images.
                    };

                    // Set the current worksheet as active.
                    workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                    // Full path for the HTML output file.
                    string htmlPath = Path.Combine(outputDir, $"{sheet.Name}.html");

                    // Save the worksheet as HTML; Aspose creates the external CSS and image files automatically.
                    workbook.Save(htmlPath, options);
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {exSheet.Message}");
                }
            }

            Console.WriteLine("Conversion completed successfully.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
