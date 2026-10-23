// Title: Create a PDF from an Excel workbook with a custom logo cover page using Aspose.Cells for .NET
// Common Searches: how to add a logo as a cover page when converting Excel to PDF with Aspose.Cells C# | Aspose.Cells C# generate PDF with a separate cover sheet containing an image | prepend a custom cover sheet to an Excel workbook before PDF export using Aspose.Cells | C# combine cover page and workbook worksheets into a single PDF with Aspose.Cells | add picture to first worksheet and export whole workbook to PDF Aspose.Cells .NET
// Tags: cover sheet creation with Aspose.Cells | logo image insertion into Excel worksheet | worksheet copy operation Aspose.Cells | PDF conversion with combined workbook Aspose.Cells | C# Aspose.Cells custom PDF cover generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel file, creates a new workbook with a first sheet that serves as a cover page, inserts a logo image and title on that sheet, copies all original worksheets into the new workbook, and then saves the combined workbook as a single PDF document.
class Program
{
    static void Main()
    {
        try
        {
            // Verify required files exist
            const string inputPath = "input.xlsx";
            const string logoPath = "logo.png";

            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            if (!File.Exists(logoPath))
                throw new FileNotFoundException($"Logo image not found: {logoPath}");

            // Load the original workbook that will become the main content of the PDF
            Workbook sourceWorkbook = new Workbook(inputPath);

            // Create a new workbook that will hold the cover page and the original sheets
            Workbook finalWorkbook = new Workbook();

            // The first (and only) sheet in the new workbook will be used as the cover page
            Worksheet coverSheet = finalWorkbook.Worksheets[0];
            coverSheet.Name = "Cover";

            // Insert the custom logo image onto the cover sheet
            int pictureIndex = coverSheet.Pictures.Add(0, 0, logoPath);
            Picture logo = coverSheet.Pictures[pictureIndex];
            logo.Width = 300;
            logo.Height = 150;

            // Optionally add a title or any other text to the cover page
            Cells cells = coverSheet.Cells;
            cells["A5"].PutValue("Report Title");
            Style titleStyle = cells["A5"].GetStyle();
            titleStyle.Font.Size = 24;
            titleStyle.Font.IsBold = true;
            cells["A5"].SetStyle(titleStyle);

            // Copy all worksheets from the source workbook into the final workbook,
            // placing them after the cover sheet
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                // Add a copy of the worksheet to the final workbook with a new name
                // Using array overload to match the API version
                finalWorkbook.Worksheets.AddCopy(new Worksheet[] { ws }, new string[] { ws.Name + "_Copy" });
            }

            // Save the combined workbook as a single PDF file
            finalWorkbook.Save("output.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
