// Title: Insert a PNG logo at (100,200) coordinates when converting an Excel workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Add a free‑floating picture to the first worksheet, set its Left property to 100 and Top property to 200, then save the workbook as PDF. | Load an Excel file, embed a PNG image, position it using absolute point coordinates, and export the sheet to PDF using the Aspose.Cells C# API.
// Common Searches: how to set picture left and top values in Aspose.Cells before PDF conversion C# | Aspose.Cells place PNG logo at specific point coordinates in generated PDF | free floating image positioning in Excel to PDF conversion using Aspose.Cells .NET
// Tags: Aspose.Cells free-floating picture positioning | C# set picture left top properties Aspose.Cells | Aspose.Cells Excel to PDF conversion with image overlay | absolute coordinate image placement Aspose.Cells | insert PNG logo into PDF using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, adds a PNG logo as a free‑floating picture on the first worksheet, positions it at 100 points from the left and 200 points from the top, and saves the result as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string logoPath = "logo.png";
            const string outputPath = "output.pdf";

            // Verify that the required files exist to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");
            if (!File.Exists(logoPath))
                throw new FileNotFoundException($"Logo image not found: {logoPath}");

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet where the logo will be placed
            Worksheet sheet = workbook.Worksheets[0];

            // Add the logo image; the method returns the index of the new picture
            int pictureIndex = sheet.Pictures.Add(0, 0, logoPath);
            Picture logo = sheet.Pictures[pictureIndex];

            // Make the picture free‑floating so it can be positioned by absolute coordinates
            logo.Placement = PlacementType.FreeFloating;

            // Position the logo using absolute coordinates (points)
            logo.Left = 100; // horizontal offset
            logo.Top = 200;  // vertical offset

            // Save the workbook as a PDF file with the logo positioned as specified
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
