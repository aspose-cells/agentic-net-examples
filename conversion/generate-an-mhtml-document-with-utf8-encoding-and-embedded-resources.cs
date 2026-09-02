// Title: Create an MHTML file with UTF-8 encoding and embedded Base64 images using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that builds a Workbook, inserts text and an optional PNG picture, configures HtmlSaveOptions for MHTML with UTF-8 encoding, Base64 image embedding, and PresentationPreference enabled, then saves the file. | Adapt an existing Aspose.Cells example to export a worksheet as an MHTML document with embedded resources while ensuring the output uses the UTF-8 character set.
// Common Searches: how to save an Excel workbook as MHTML with embedded images in C# using Aspose.Cells | Aspose.Cells C# export to MHTML UTF-8 encoding and base64 pictures | set PresentationPreference true when converting Excel to MHTML with Aspose.Cells | generate MHTML from workbook with optional logo image Aspose.Cells example
// Tags: Aspose.Cells MHTML export base64 images | C# HtmlSaveOptions UTF-8 MHTML | Aspose.Cells PresentationPreference enabled | embed PNG picture in MHTML Aspose.Cells | create MHTML document from Excel workbook C#

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample creates a new Workbook, writes a text value, optionally adds a PNG picture, configures HtmlSaveOptions for MHTML with UTF-8 encoding, Base64 image embedding, and PresentationPreference set to true, and saves the result as an MHTML file (output.mht).
    class GenerateMhtml
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Hello, MHTML!");

                // Add an image if the file exists
                const string imagePath = "logo.png";
                if (File.Exists(imagePath))
                {
                    // Add picture at row 2, column 2 (zero‑based indexing)
                    worksheet.Pictures.Add(2, 2, imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Configure HTML save options for MHTML format
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml)
                {
                    Encoding = Encoding.UTF8,
                    ExportImagesAsBase64 = true,
                    PresentationPreference = true
                };

                // Save the workbook as an MHTML document with embedded resources
                const string outputPath = "output.mht";
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
