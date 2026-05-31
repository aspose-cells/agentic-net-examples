using System;
using System.IO;
using Aspose.Cells;

class HtmlToPdfWithCustomFont
{
    static void Main()
    {
        try
        {
            // Specify the folder that contains custom TrueType fonts (recursive scan enabled).
            FontConfigs.SetFontFolder(@"C:\CustomFonts", true);

            // Define font substitutes for missing fonts.
            FontConfigs.SetFontSubstitutes("Times New Roman", new[] { "Liberation Serif", "DejaVu Serif" });

            string inputPath = @"C:\Input\sample.html";
            string outputPath = @"C:\Output\sample.pdf";

            // Prevent FileNotFoundException for the input HTML file.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the source HTML file into a Workbook.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DefaultFont = "Arial",               // Fallback font.
                CheckWorkbookDefaultFont = true,     // Use workbook's default font first.
                CheckFontCompatibility = true        // Verify font compatibility per character.
            };

            // Save the workbook as a PDF file using the configured options.
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}