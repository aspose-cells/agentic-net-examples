// Title: Insert an external SVG file into cells E5:F6 of an Excel worksheet while preserving its original size using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to embed an SVG file at cell E5 without resizing | Maintain original SVG image size when adding to Excel via Aspose.Cells | Place a free‑floating SVG picture in a worksheet range using Aspose.Cells .NET | Programmatically add logo.svg to an Excel workbook with Aspose.Cells | Set picture placement to free floating and keep dimensions in Aspose.Cells
// Tags: Aspose.Cells add SVG picture to worksheet | preserve original SVG dimensions Aspose.Cells | free floating picture placement Aspose.Cells | C# insert image at specific cell range Aspose.Cells | set picture width height to original size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example verifies the presence of logo.svg, creates a new workbook, and inserts the SVG as a picture anchored at cell E5 (the top‑left corner of range E5:F6). The picture is set to free‑floating placement and its Width and Height are assigned the OriginalWidth and OriginalHeight values to retain the SVG's native dimensions before saving the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Path to the SVG file that will be inserted.
        const string svgPath = "logo.svg";

        // Verify that the SVG file exists.
        if (!File.Exists(svgPath))
        {
            Console.WriteLine($"Error: The file '{svgPath}' was not found.");
            return;
        }

        // Create a new workbook and get the first worksheet.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Target cell range: E5:F6.
        // Row and column indices are zero‑based, so E5 corresponds to row 4, column 4.
        int targetRow = 4;    // Row 5 in Excel (E5)
        int targetColumn = 4; // Column E in Excel

        // Add the SVG picture to the worksheet.
        // The Add method places the picture with its top‑left corner at the specified cell.
        int pictureIndex = sheet.Pictures.Add(targetRow, targetColumn, svgPath);
        Picture picture = sheet.Pictures[pictureIndex];

        // Ensure the picture is free‑floating so it does not move with cells.
        picture.Placement = PlacementType.FreeFloating;

        // Preserve the original dimensions of the SVG.
        // OriginalWidth and OriginalHeight return the size of the image in pixels.
        picture.Width = picture.OriginalWidth;
        picture.Height = picture.OriginalHeight;

        // Optionally, you can adjust the picture's position to align exactly with the range.
        // The range E5:F6 spans two columns and two rows. We keep the top‑left corner at E5.
        // No scaling is performed, so the picture may extend beyond F6 if it is larger.

        // Save the workbook.
        const string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
        Console.WriteLine($"SVG inserted and workbook saved to '{outputPath}'.");
    }
}
