using System;
using System.IO;
using Aspose.Cells;

class ReplaceWorksheetBackground
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string backgroundPath = "newBackground.jpg";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the target worksheet (first sheet in this example)
            Worksheet worksheet = workbook.Worksheets[0];

            // Preserve current page‑setup settings
            PageSetup pageSetup = worksheet.PageSetup;
            var orientation   = pageSetup.Orientation;
            var leftMargin    = pageSetup.LeftMargin;
            var rightMargin   = pageSetup.RightMargin;
            var topMargin     = pageSetup.TopMargin;
            var bottomMargin  = pageSetup.BottomMargin;
            var paperSize     = pageSetup.PaperSize;

            // Replace the worksheet background image if the file exists
            if (File.Exists(backgroundPath))
            {
                byte[] newBackground = File.ReadAllBytes(backgroundPath);
                worksheet.BackgroundImage = newBackground;
            }
            else
            {
                Console.WriteLine($"Background image not found: {backgroundPath}. Skipping background replacement.");
            }

            // Re‑apply the preserved page‑setup settings (ensures they remain unchanged)
            pageSetup.Orientation   = orientation;
            pageSetup.LeftMargin    = leftMargin;
            pageSetup.RightMargin   = rightMargin;
            pageSetup.TopMargin     = topMargin;
            pageSetup.BottomMargin  = bottomMargin;
            pageSetup.PaperSize     = paperSize;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}