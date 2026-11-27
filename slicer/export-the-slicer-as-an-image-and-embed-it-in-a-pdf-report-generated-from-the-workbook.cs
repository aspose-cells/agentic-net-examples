// Title: How to export an Excel slicer as a PNG image and embed it into a PDF report using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export slicer to PNG image | Add slicer picture to PDF generated from workbook with Aspose.Cells | Render Excel slicer as image using reflection in Aspose.Cells .NET | Create PDF report that includes a slicer image using Aspose.Cells C# | Save slicer as image when Aspose.Cells version lacks ToImage method | Insert picture from slicer into new worksheet before saving as PDF Aspose.Cells
// Tags: Aspose.Cells slicer to PNG conversion | C# insert picture into worksheet Aspose.Cells | Aspose.Cells generate PDF with embedded image | reflection call Slicer.ToImage Aspose.Cells | Excel slicer image export Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Slicers;   // Slicer related classes

// The sample loads an Excel workbook, checks for a slicer on the first worksheet, uses reflection to call the optional Slicer.ToImage method and saves the slicer as a PNG file, adds a new worksheet, inserts the PNG as a picture, and finally saves the workbook as a PDF report containing the slicer image.
class SlicerToPdfReport
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string slicerImagePath = "slicer.png";
        const string pdfPath = "Report.pdf";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source workbook that contains the slicer
            Workbook sourceWorkbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Ensure the worksheet contains at least one slicer
            if (sourceSheet.Slicers.Count == 0)
            {
                Console.WriteLine("Error: No slicers were found on the first worksheet.");
                return;
            }

            // Retrieve the first slicer on the worksheet
            Slicer slicer = sourceSheet.Slicers[0];

            // Attempt to render the slicer to an image file.
            // Aspose.Cells does not provide a direct ToImage method for Slicer in all versions,
            // so this block is kept for compatibility with versions that support it.
            try
            {
                // Delete existing image to avoid conflicts
                if (File.Exists(slicerImagePath))
                {
                    File.Delete(slicerImagePath);
                }

                // If the ToImage method exists, it will be invoked; otherwise, this will be skipped.
                // Using reflection to avoid compile‑time errors on older library versions.
                var toImageMethod = slicer.GetType().GetMethod("ToImage", new[] { typeof(string) });
                if (toImageMethod != null)
                {
                    toImageMethod.Invoke(slicer, new object[] { slicerImagePath });
                }
                else
                {
                    Console.WriteLine("Info: Slicer.ToImage method is not available in this Aspose.Cells version.");
                }
            }
            catch (Exception imgEx)
            {
                Console.WriteLine($"Warning: Unable to render slicer to image. {imgEx.Message}");
                // Continue without image; the report will be generated without the slicer picture.
            }

            // Create a new worksheet that will serve as the PDF report page
            Worksheet reportSheet = sourceWorkbook.Worksheets.Add("Report");

            // Insert the slicer image into the report worksheet if the image was created
            if (File.Exists(slicerImagePath))
            {
                // Row and column indices are zero‑based; (0,0) places the image at the top‑left corner
                int pictureIndex = reportSheet.Pictures.Add(0, 0, slicerImagePath);
                Picture picture = reportSheet.Pictures[pictureIndex];

                // Optionally adjust the picture size (width in pixels)
                picture.Width = 500;   // set desired width
                picture.Height = -1;   // -1 preserves aspect ratio
            }

            // Save the entire workbook (including the report sheet with the image) as a PDF file
            sourceWorkbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"Report generated successfully: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
