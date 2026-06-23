using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfBackgroundImageHandler
{
    static void Main()
    {
        try
        {
            // Configurable size threshold (e.g., 5 MB)
            const long sizeThreshold = 5 * 1024 * 1024;

            // ---------- Create ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            for (int i = 0; i < 100; i++)
            {
                sheet.Cells[i, 0].PutValue($"Row {i + 1}");
            }

            // Add a background image (if the file exists)
            string bgImagePath = "background.jpg";
            if (File.Exists(bgImagePath))
            {
                // Adding picture at the top‑left corner; this simulates a background image
                sheet.Pictures.Add(0, 0, bgImagePath);
            }

            // ---------- First Save (in‑memory) ----------
            PdfSaveOptions firstOptions = new PdfSaveOptions
            {
                OptimizationType = PdfOptimizationType.Standard
            };

            using (MemoryStream tempStream = new MemoryStream())
            {
                workbook.Save(tempStream, firstOptions);
                long pdfSize = tempStream.Length;

                // ---------- Conditional Logic ----------
                if (pdfSize > sizeThreshold)
                {
                    // Size exceeds threshold → disable background images and apply stronger optimization
                    PdfSaveOptions reducedOptions = new PdfSaveOptions
                    {
                        OptimizationType = PdfOptimizationType.MinimumSize
                    };

                    // Make background transparent (effectively removes background images)
                    // Note: ImageOrPrintOptions may not be available in older versions; this line is optional.
                    // reducedOptions.ImageOrPrintOptions.Transparent = true;

                    // Optional: downsample images to further shrink the file
                    reducedOptions.SetImageResample(96, 70); // 96 PPI, 70 % JPEG quality

                    // Save the reduced PDF
                    workbook.Save("output_reduced.pdf", reducedOptions);
                    Console.WriteLine($"PDF size {pdfSize} bytes exceeded threshold. Saved reduced PDF.");
                }
                else
                {
                    // Size within limit → keep original PDF
                    workbook.Save("output_original.pdf", firstOptions);
                    Console.WriteLine($"PDF size {pdfSize} bytes within threshold. Saved original PDF.");
                }
            }
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File not found: {ex.FileName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}