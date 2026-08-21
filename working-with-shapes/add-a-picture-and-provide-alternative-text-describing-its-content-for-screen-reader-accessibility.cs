// Title: Insert a Picture with Alt Text into an Excel Worksheet using Aspose.Cells for .NET
// Description: Demonstrates how to create a new Workbook, add a PNG image to cell C3 (row 2, column 2), assign descriptive AlternativeText for screen‑reader accessibility, handle missing files gracefully, and save the file as WorkbookWithPicture.xlsx.
// Keywords: Aspose.Cells | .NET | C# | insert picture Excel | add image worksheet | AlternativeText property | alt text accessibility | Excel picture insertion | screen reader support | Aspose.Cells picture.Add
// Common Searches: Aspose.Cells add image to Excel with alt text | C# set picture AlternativeText in Aspose.Cells | how to make Excel images accessible using Aspose | insert picture into worksheet Aspose.Cells .NET | alternative text for Excel picture Aspose
// Developer Intent: Add an image to a worksheet and provide a descriptive alt text to meet accessibility requirements.
// Use Cases: Embed a company logo in generated reports with compliance‑ready alt text. | Add product thumbnails to a catalog sheet while ensuring screen‑reader descriptions. | Programmatically insert diagrams into exported workbooks and supply accessibility metadata.
// AI Prompts: Generate C# code that uses Aspose.Cells to place a PNG at cell C3 and set its AlternativeText to a custom description. | Create a reusable method that checks for an image file, inserts it into a worksheet, logs a warning if absent, and applies alt text. | Write a script to batch‑insert multiple pictures across different worksheets, assigning unique AlternativeText to each.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsAddPictureWithAltText
{
    // Demonstrates how to create a new Workbook, add a PNG image to cell C3 (row 2, column 2), assign descriptive AlternativeText for screen‑reader accessibility, handle missing files gracefully, and save the file as WorkbookWithPicture.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the picture file to be inserted
                string picturePath = "image.png";

                if (File.Exists(picturePath))
                {
                    // Add the picture at row 2, column 2 (zero‑based indices)
                    int pictureIndex = worksheet.Pictures.Add(2, 2, picturePath);
                    Picture picture = worksheet.Pictures[pictureIndex];

                    // Set alternative (alt) text for screen readers
                    picture.AlternativeText = "Company logo showing a blue circle with white initials";
                }
                else
                {
                    Console.WriteLine($"Picture file not found: {picturePath}. Skipping picture insertion.");
                }

                // Save the workbook to a file
                string outputPath = "WorkbookWithPicture.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
