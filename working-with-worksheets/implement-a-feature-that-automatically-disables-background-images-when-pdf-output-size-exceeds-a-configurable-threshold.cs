// Title: How to automatically disable worksheet background images in Aspose.Cells for .NET when the exported PDF exceeds a size limit
// Common Searches: Aspose.Cells C# disable worksheet background images if PDF size exceeds 5 MB | Check PDF file size before saving with Aspose.Cells and remove background images | Conditional PDF export in Aspose.Cells based on output size threshold | Programmatically clear Excel worksheet background images during PDF export using Aspose.Cells | C# Aspose.Cells export to PDF with size constraint and background image removal
// Tags: size-based PDF export decision Aspose.Cells | programmatic worksheet background image clearing | pdf size evaluation via memory stream Aspose.Cells | large PDF export background image handling C# | asp.net cells pdf output size limit strategy

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    // The program loads an Excel workbook, saves it to a memory stream as PDF, checks the generated PDF size against a configurable byte limit, and if the size exceeds the limit it clears the BackgroundImage property on every worksheet before saving the final PDF; otherwise it writes the original PDF stream directly to the output file.
    class Program
    {
        static void Main(string[] args)
        {
            // Configurable parameters
            string inputFilePath = "input.xlsx";          // Path to the source workbook
            string outputFilePath = "output.pdf";         // Desired PDF output path
            long sizeThresholdBytes = 5 * 1024 * 1024;    // Threshold (e.g., 5 MB)

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(inputFilePath);

            // Save to a memory stream first to evaluate the resulting PDF size
            using (MemoryStream tempStream = new MemoryStream())
            {
                // Save as PDF to the stream (uses the provided save rule)
                workbook.Save(tempStream, SaveFormat.Pdf);

                // If the generated PDF exceeds the threshold, remove background images
                if (tempStream.Length > sizeThresholdBytes)
                {
                    // Disable background images on all worksheets
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        sheet.BackgroundImage = null;
                    }

                    // Save the final PDF after background images have been disabled
                    workbook.Save(outputFilePath, SaveFormat.Pdf);
                }
                else
                {
                    // PDF size is within the limit; write the stream directly to the output file
                    File.WriteAllBytes(outputFilePath, tempStream.ToArray());
                }
            }
        }
    }
}
