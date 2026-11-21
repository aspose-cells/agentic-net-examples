// Title: Add a digital signature image with caption to every worksheet in an Excel file using Aspose.Cells for .NET
// Common Searches: how to insert a signature PNG into each sheet with Aspose.Cells C# | Aspose.Cells C# add picture and textbox at specific row and column | C# loop through all worksheets and place a logo at the bottom | set picture size and add centered text box above it using Aspose.Cells | save modified workbook to new file after adding signature in C#
// Tags: Aspose.Cells embed image in worksheet | Aspose.Cells create caption textbox | C# define picture width height Aspose.Cells | Aspose.Cells process each worksheet | C# save workbook after adding signature image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// The example loads an Excel workbook, iterates over every worksheet, embeds a PNG signature image anchored at a specific row and column, adjusts its width and height, adds a centered textbox caption above the image, and finally saves the updated workbook to a new file while handling missing files and directory creation.
class ExcelSignatureOverlay
{
    static void Main()
    {
        // Paths for the source Excel file, signature image, and the output Excel file
        string sourceExcelPath = @"C:\Docs\input.xlsx";
        string signatureImagePath = @"C:\Images\signature.png";
        string outputExcelPath = @"C:\Docs\output_signed.xlsx";

        try
        {
            // Verify that required files exist
            if (!File.Exists(sourceExcelPath))
                throw new FileNotFoundException("Source Excel file not found.", sourceExcelPath);
            if (!File.Exists(signatureImagePath))
                throw new FileNotFoundException("Signature image file not found.", signatureImagePath);

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputExcelPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the workbook
            Workbook workbook = new Workbook(sourceExcelPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Add the signature image to the worksheet (anchor at row 20, column 0)
                    using (FileStream imgStream = new FileStream(signatureImagePath, FileMode.Open, FileAccess.Read))
                    {
                        int pictureIndex = sheet.Pictures.Add(20, 0, imgStream);
                        Picture picture = sheet.Pictures[pictureIndex];
                        picture.Width = 150;   // pixels
                        picture.Height = 50;   // pixels
                    }

                    // Add a textbox shape above the signature image (anchor at row 19, column 0)
                    Shape textShape = sheet.Shapes.AddTextBox(19, 0, 0, 0, 150, 20);
                    textShape.Text = "Digitally signed by John Doe";
                    textShape.TextHorizontalAlignment = TextAlignmentType.Center;
                    textShape.TextVerticalAlignment = TextAlignmentType.Center;

                    // Apply font styling
                    Font font = textShape.Font;
                    font.Size = 12;
                    font.IsBold = true;
                    font.Color = Color.Black;
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {exSheet.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputExcelPath);
            Console.WriteLine($"Workbook saved successfully to '{outputExcelPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
