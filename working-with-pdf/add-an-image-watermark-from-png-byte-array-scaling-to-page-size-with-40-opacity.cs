// Title: Add a PNG watermark from a byte array to an Aspose.Cells worksheet and scale it to the page size in C#
// Common Searches: C# Aspose.Cells add PNG watermark from stream and fit to worksheet page | scale image watermark to cover entire Excel sheet using Aspose.Cells | estimate worksheet page size from column and row pixel dimensions in Aspose.Cells | set picture position top left and resize to page size with Aspose.Cells C# | load PNG into Aspose.Cells picture without using IsBackground property
// Tags: Aspose.Cells picture watermark insertion | worksheet page size estimation in pixels | resize picture to match worksheet area | place picture at origin of worksheet | save workbook with embedded watermark

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a PNG file into a stream, creates a new Workbook, adds the image as a picture to the first worksheet, estimates the worksheet page size using column widths and row heights, resizes the picture to cover that area, positions it at the top‑left corner, and saves the workbook; opacity handling is omitted for compatibility.
class AddWatermarkExample
{
    static void Main()
    {
        try
        {
            const string watermarkFile = "watermark.png";

            // Verify that the watermark image exists
            if (!File.Exists(watermarkFile))
            {
                Console.WriteLine($"Watermark image not found: {watermarkFile}");
                return;
            }

            // Load the image into a stream (required by Pictures.Add)
            using (FileStream imgStream = File.OpenRead(watermarkFile))
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add the picture to the worksheet (returns the picture index)
                int pictureIndex = sheet.Pictures.Add(0, 0, imgStream);
                Picture picture = sheet.Pictures[pictureIndex];

                // NOTE: In some Aspose.Cells versions the IsBackground and Transparency
                // properties are not available. They are omitted here to ensure compatibility.

                // Estimate page size in pixels using the default column/row dimensions
                int totalColumns = sheet.Cells.MaxColumn + 1; // include the last column
                int totalRows = sheet.Cells.MaxRow + 1;       // include the last row

                // Fallback to a reasonable default if the sheet is empty
                if (totalColumns == 0) totalColumns = 10;
                if (totalRows == 0) totalRows = 30;

                int columnWidthPx = sheet.Cells.GetColumnWidthPixel(0);
                int rowHeightPx = sheet.Cells.GetRowHeightPixel(0);

                int pageWidthPixels = columnWidthPx * totalColumns;
                int pageHeightPixels = rowHeightPx * totalRows;

                // Resize the picture to cover the estimated page area
                picture.Width = pageWidthPixels;
                picture.Height = pageHeightPixels;

                // Position the watermark at the top‑left corner
                picture.Left = 0;
                picture.Top = 0;

                // Save the workbook
                const string outputFile = "WorkbookWithWatermark.xlsx";
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to {outputFile}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
