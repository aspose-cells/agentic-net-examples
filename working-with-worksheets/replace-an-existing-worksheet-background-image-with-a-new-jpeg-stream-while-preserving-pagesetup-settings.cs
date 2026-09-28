// Title: Replace an Excel worksheet’s background image with a JPEG byte array while keeping all PageSetup settings using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads a JPEG file into a byte[] and assigns it to Worksheet.BackgroundImage, then restores the original PageSetup properties. | Demonstrate removing the current worksheet background, setting a new JPEG from a stream, and reapplying orientation, margins, and fit‑to‑page settings with Aspose.Cells.
// Common Searches: Aspose.Cells how to change worksheet background image without affecting page margins | C# replace Excel sheet background with JPEG byte array preserving page setup | set worksheet background image from stream Aspose.Cells .NET example | keep orientation and margins after updating worksheet background in Aspose.Cells
// Tags: replace worksheet background Aspose.Cells | assign background from byte array C# | retain page setup after background change | load JPEG into Excel worksheet Aspose.Cells | remove current worksheet background prior to new image

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // The example loads a workbook, saves the first worksheet's PageSetup values, clears any existing background, reads a JPEG file into a byte array, assigns it to Worksheet.BackgroundImage, restores the saved PageSetup properties, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "InputWorkbook.xlsx";
                const string backgroundPath = "NewBackground.jpg";
                const string outputPath = "OutputWorkbook.xlsx";

                // Verify input files exist
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input workbook not found: {inputPath}");
                    return;
                }

                if (!File.Exists(backgroundPath))
                {
                    Console.WriteLine($"Background image not found: {backgroundPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Preserve current page‑setup settings
                PageSetup ps = sheet.PageSetup;
                var orientation = ps.Orientation;
                var paperSize = ps.PaperSize;
                var topMargin = ps.TopMargin;
                var bottomMargin = ps.BottomMargin;
                var leftMargin = ps.LeftMargin;
                var rightMargin = ps.RightMargin;
                var headerMargin = ps.HeaderMargin;
                var footerMargin = ps.FooterMargin;
                var fitToPagesTall = ps.FitToPagesTall;
                var fitToPagesWide = ps.FitToPagesWide;

                // ---------- Replace the background image ----------
                // Clear any existing background image
                sheet.BackgroundImage = null;

                // Load the new JPEG image into a byte array and assign it
                try
                {
                    byte[] imageBytes;
                    using (FileStream fs = File.OpenRead(backgroundPath))
                    {
                        imageBytes = new byte[fs.Length];
                        int bytesRead = fs.Read(imageBytes, 0, imageBytes.Length);
                        if (bytesRead != imageBytes.Length)
                        {
                            Console.WriteLine("Failed to read the entire background image file.");
                            return;
                        }
                    }
                    sheet.BackgroundImage = imageBytes;
                }
                catch (Exception imgEx)
                {
                    Console.WriteLine($"Error loading background image: {imgEx.Message}");
                    return;
                }

                // ---------- Re‑apply the preserved page‑setup settings ----------
                ps.Orientation = orientation;
                ps.PaperSize = paperSize;
                ps.TopMargin = topMargin;
                ps.BottomMargin = bottomMargin;
                ps.LeftMargin = leftMargin;
                ps.RightMargin = rightMargin;
                ps.HeaderMargin = headerMargin;
                ps.FooterMargin = footerMargin;
                ps.FitToPagesTall = fitToPagesTall;
                ps.FitToPagesWide = fitToPagesWide;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
