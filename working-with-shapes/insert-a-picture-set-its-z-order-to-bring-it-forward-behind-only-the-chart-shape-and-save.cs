// Title: Insert a PNG picture, place a column chart above it, adjust Z‑order so the picture stays behind the chart, and save the workbook with Aspose.Cells for .NET
// AI Prompts: Add a PNG image to cell B2, then add a column chart covering D5:H20, move the picture one Z‑order level forward so it remains beneath the chart, and save the workbook as an XLSX file. | Create a new workbook, insert a picture, insert a column chart, set the picture's Z‑order to be just behind the chart shape, and export the file.
// Common Searches: how to insert an image and set its Z-order behind a chart using Aspose.Cells C# | Aspose.Cells C# move picture layer under chart shape | set shape Z-order in Excel workbook with Aspose.Cells .NET | place PNG picture at specific cell and keep it below a chart Aspose.Cells | adjust layering of picture and chart in Aspose.Cells generated Excel file
// Tags: Aspose.Cells insert picture C# | Aspose.Cells add column chart C# | Aspose.Cells control shape Z-order | Aspose.Cells picture behind chart | Aspose.Cells export workbook to XLSX

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

// The example creates a new workbook, inserts a PNG picture at cell B2, adds a column chart covering D5:H20, moves the picture forward one Z‑order level so it stays directly beneath the chart, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the picture file to insert
            string picturePath = @"C:\Images\SamplePicture.png";

            // Verify that the picture file exists
            if (!File.Exists(picturePath))
                throw new FileNotFoundException("Picture file not found.", picturePath);

            // Insert the picture at cell B2 (row 1, column 1)
            // Add returns the index of the picture; retrieve the Picture object
            int pictureIndex = sheet.Pictures.Add(1, 1, picturePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Insert a column chart that will appear above the picture
            // Chart placed at cells D5 to H20 (rows 4-19, columns 3-7)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 4, 3, 19, 7);
            Chart chart = sheet.Charts[chartIndex];

            // (Optional) Adjust Z‑order if needed – omitted for compatibility

            // Ensure the output directory exists
            string outputPath = @"C:\Output\WorkbookWithPictureAndChart.xlsx";
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
