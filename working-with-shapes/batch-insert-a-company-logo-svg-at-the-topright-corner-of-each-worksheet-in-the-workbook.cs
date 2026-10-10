// Title: Add a company logo SVG to the top‑right corner of every worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Place an SVG image at the first row of the right‑most used column on each worksheet, set its width to 150 px and height to 50 px, then save the workbook with Aspose.Cells. | Loop through all worksheets, insert a picture from a file path into the top‑right cell, adjust the picture size, and persist the updated file.
// Common Searches: Aspose.Cells C# add same SVG logo to every sheet in a workbook | how to position an image in the top right corner of all worksheets using Aspose.Cells | batch insert company logo into multiple Excel worksheets Aspose.Cells .NET | set picture dimensions when adding an SVG to a worksheet with Aspose.Cells | determine rightmost used column to anchor a picture in Aspose.Cells
// Tags: add svg picture to each worksheet Aspose.Cells | position logo at top right cell Aspose.Cells | resize picture dimensions Aspose.Cells C# | iterate worksheets batch insert image Aspose.Cells | use Pictures.Add with SVG Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, checks that both the workbook and the SVG logo file exist, then iterates through every worksheet. For each sheet it finds the right‑most used column, adds the SVG logo anchored to the first row of that column, resizes the picture to 150 × 50 pixels, and finally saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Input and output file paths
                string inputPath = "InputWorkbook.xlsx";
                string logoPath = "CompanyLogo.svg";
                string outputPath = "OutputWorkbook.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input workbook not found at '{inputPath}'.");
                    return;
                }

                // Verify that the logo file exists
                if (!File.Exists(logoPath))
                {
                    Console.WriteLine($"Error: Logo file not found at '{logoPath}'.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the right‑most used column (0‑based). If the sheet is empty, MaxColumn will be 0.
                    int rightMostColumn = sheet.Cells.MaxColumn;

                    // Add the SVG picture anchored to the first row of the right‑most column
                    int pictureIndex = sheet.Pictures.Add(0, rightMostColumn, logoPath);

                    // Adjust picture size if needed
                    Picture picture = sheet.Pictures[pictureIndex];
                    picture.Width = 150;   // desired width in pixels
                    picture.Height = 50;   // desired height in pixels
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
