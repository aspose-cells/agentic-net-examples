using System;
using System.IO;
using Aspose.Cells;

class RenderingWarningCallback : IWarningCallback
{
    public int WarningCount { get; private set; }

    // Called by Aspose.Cells when a warning occurs during rendering/saving.
    public void Warning(WarningInfo warningInfo)
    {
        // Capture only font substitution warnings.
        if (warningInfo.WarningType == WarningType.FontSubstitution)
        {
            Console.WriteLine($"Font substitution warning: {warningInfo.Description}");
            WarningCount++;
        }
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // -------------------------------------------------
            // 1. Create a workbook and add content that will trigger a font substitution warning.
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Cell with a font that does not exist on the system.
            sheet.Cells["A1"].PutValue("Text with a missing font");
            Style style = workbook.CreateStyle();
            style.Font.Name = "NonExistentFont";
            sheet.Cells["A1"].SetStyle(style);

            // -------------------------------------------------
            // 2. Configure PDF save options and attach the custom warning callback.
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Create and assign the warning callback.
            RenderingWarningCallback warningCallback = new RenderingWarningCallback();
            pdfOptions.WarningCallback = warningCallback;   // Captures rendering warnings.

            // Optional: set a default font to reduce warnings for other characters.
            pdfOptions.DefaultFont = "Arial";

            // -------------------------------------------------
            // 3. Save the workbook to PDF. Ensure the output directory exists.
            // -------------------------------------------------
            string outputPath = "output.pdf";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, pdfOptions);

            // -------------------------------------------------
            // 4. Log the total number of font substitution warnings captured.
            // -------------------------------------------------
            Console.WriteLine($"Total font substitution warnings captured: {warningCallback.WarningCount}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}