// Title: Insert a logo image at the top of an Excel worksheet and freeze the header row using Aspose.Cells for .NET (C#)
// AI Prompts: Place a PNG logo into cell A1 of the first worksheet and then apply FreezePanes to keep the first row visible with Aspose.Cells in C#. | Create a new workbook, add a picture from a file path to the top-left corner, freeze row 1, and save as XLSX using the Aspose.Cells API. | Modify the code to anchor a logo image to cell A1, optionally set its placement, and freeze the header row so the branding remains on screen.
// Common Searches: Aspose.Cells C# how to insert a logo at the top of a sheet and freeze the header row | freeze first row after adding picture to Excel with Aspose.Cells .NET | add PNG image to cell A1 and keep it visible while scrolling using Aspose.Cells | example of using FreezePanes after inserting a picture in Aspose.Cells for .NET
// Tags: add picture to worksheet Aspose.Cells C# | freeze header row Aspose.Cells | Aspose.Cells insert PNG logo | Aspose.Cells FreezePanes example | picture anchoring cell A1 Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program creates a new workbook, inserts a 'logo.png' image into cell A1 of the first worksheet (if the file exists), freezes the first row so the logo stays visible while scrolling, and saves the file as 'output.xlsx' using Aspose.Cells for .NET.
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

            // Path to the logo image file (replace with actual file path)
            string logoPath = "logo.png";

            // Insert the logo picture at cell A1 if the file exists
            if (File.Exists(logoPath))
            {
                // The picture will be anchored to the cell at row 0, column 0
                int pictureIndex = sheet.Pictures.Add(0, 0, logoPath);
                // Optional: adjust picture properties (size, placement) if needed
                // sheet.Pictures[pictureIndex].Placement = PlacementType.FreeFloating;
            }
            else
            {
                Console.WriteLine($"Warning: Logo file not found at '{logoPath}'. Skipping picture insertion.");
            }

            // Freeze the first row so the logo (if inserted) stays visible
            // FreezePanes(row, column, totalRows, totalColumns)
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
