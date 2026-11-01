// Title: Add a shared background image to the first worksheet of multiple Excel workbooks and export each workbook to PDF with Aspose.Cells for .NET
// Common Searches: how to apply the same background PNG to the first sheet of several Excel files using Aspose.Cells in C# | batch add a picture to the first worksheet of multiple workbooks and convert them to PDF with Aspose.Cells | using MemoryStream to set a free‑floating background image on worksheets in Aspose.Cells C# | process a list of Excel files, add a background image, and save as PDF programmatically with Aspose.Cells | C# code to load a PNG once and reuse it as worksheet background for many workbooks before PDF export
// Tags: add background image to first worksheet Aspose.Cells | handle several Excel files C# Aspose.Cells | export workbook as PDF Aspose.Cells | picture placement freefloating Aspose.Cells | reuse image data with MemoryStream Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program iterates over a collection of Excel files, loads a PNG image once into a byte array, inserts it as a free‑floating picture anchored at cell A1 of each workbook's first worksheet via a MemoryStream, and then saves each workbook as a PDF.
class Program
{
    static void Main()
    {
        // List of workbook files to process
        var workbookPaths = new List<string>
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx"
            // add additional workbook file names here
        };

        // Path to the background image file
        const string backgroundImageFile = "Background.png";

        // Verify that the background image exists
        if (!File.Exists(backgroundImageFile))
        {
            Console.WriteLine($"Background image file not found: {backgroundImageFile}");
            return;
        }

        // Load image data once
        byte[] imageData;
        try
        {
            imageData = File.ReadAllBytes(backgroundImageFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to read background image: {ex.Message}");
            return;
        }

        foreach (string wbPath in workbookPaths)
        {
            // Verify workbook file exists
            if (!File.Exists(wbPath))
            {
                Console.WriteLine($"Workbook file not found: {wbPath}");
                continue;
            }

            try
            {
                // Load the workbook
                var workbook = new Workbook(wbPath);

                // Access the first worksheet
                var firstSheet = workbook.Worksheets[0];

                // Add the background image as a picture anchored at cell A1
                using (var ms = new MemoryStream(imageData))
                {
                    int pictureIndex = firstSheet.Pictures.Add(0, 0, ms);
                    Picture picture = firstSheet.Pictures[pictureIndex];

                    // Set the picture to free-floating so it behaves like a background
                    picture.Placement = PlacementType.FreeFloating;
                }

                // Define the output PDF file name (same name, .pdf extension)
                string pdfPath = Path.ChangeExtension(wbPath, ".pdf");

                // Export the workbook to PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"Processed '{wbPath}' -> '{pdfPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{wbPath}': {ex.Message}");
            }
        }
    }
}
