// Title: Aspose.Cells C# – Export Excel Sparklines as Inline Images in PDF
// Description: Load an Excel workbook that contains sparkline groups, render each sparkline to a high‑resolution PNG, replace the sparkline with an inline picture in the same cell, clear the original groups, and save the sheet as a PDF so the sparklines appear as embedded graphics.
// Keywords: Aspose.Cells sparkline to PDF | C# export sparkline as image | inline picture Excel PDF | render sparkline PNG Aspose | convert sparkline group PDF | high resolution sparkline image | Aspose.Cells PDF export
// Common Searches: how to export sparklines to PDF with Aspose.Cells | replace Excel sparkline with image before PDF conversion C# | Aspose.Cells render sparkline as PNG | inline picture for sparkline Aspose.Cells .NET | save workbook with sparklines as PDF
// Developer Intent: Transform an Excel file that contains sparklines into a PDF where each sparkline is displayed as an inline image rather than a dynamic chart.
// Use Cases: Produce printable financial statements that retain sparkline trends as static graphics. | Generate dashboard PDFs from Excel templates where sparklines must render consistently on all devices. | Automate archival of Excel reports with sparklines, ensuring visual fidelity in PDF format.
// AI Prompts: Create C# code with Aspose.Cells that converts every sparkline in a worksheet to a PNG and embeds it as an inline picture before saving as PDF. | Explain how to set ImageOrPrintOptions resolution for sparkline images to achieve high‑quality PDF output. | Guide me on processing multiple worksheets and sparkline groups when replacing sparklines with images for PDF export.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace SparklinePdfExport
{
    // Load an Excel workbook that contains sparkline groups, render each sparkline to a high‑resolution PNG, replace the sparkline with an inline picture in the same cell, clear the original groups, and save the sheet as a PDF so the sparklines appear as embedded graphics.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input workbook that should contain sparklines
                string inputFile = "input_with_sparklines.xlsx";

                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file \"{inputFile}\" not found. Please ensure the file exists in the application directory.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Options for rendering sparklines to PNG images
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                // Process each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Process each sparkline group
                    foreach (SparklineGroup group in sheet.SparklineGroups)
                    {
                        // Convert each sparkline to an image and insert it back into the cell
                        for (int i = 0; i < group.Sparklines.Count; i++)
                        {
                            Sparkline sparkline = group.Sparklines[i];

                            using (MemoryStream imgStream = new MemoryStream())
                            {
                                sparkline.ToImage(imgStream, imgOptions);
                                imgStream.Position = 0; // Reset stream before adding to picture collection

                                // Insert the image as an inline picture at the sparkline's cell location
                                sheet.Pictures.Add(sparkline.Row, sparkline.Column, imgStream);
                            }
                        }
                    }

                    // Remove original sparklines so they don't appear twice in the PDF
                    sheet.SparklineGroups.Clear();
                }

                // Save the result as PDF
                string outputPdf = "output_sparklines.pdf";
                workbook.Save(outputPdf, SaveFormat.Pdf);

                Console.WriteLine("Workbook with sparklines exported to PDF successfully.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
