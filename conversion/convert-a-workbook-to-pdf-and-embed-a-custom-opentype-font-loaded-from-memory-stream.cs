using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace MyApp
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string fontPath = "MyCustomFont.otf";

                // Load and register the custom font if the file exists
                if (File.Exists(fontPath))
                {
                    byte[] fontData = File.ReadAllBytes(fontPath);
                    MemoryFontSource memoryFont = new MemoryFontSource(fontData);
                    FontSourceBase[] fontSources = new FontSourceBase[] { memoryFont };
                    FontConfigs.SetFontSources(fontSources);
                }
                else
                {
                    Console.WriteLine($"Font file '{fontPath}' not found. Continuing with default fonts.");
                }

                // Create a new workbook and add sample text
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Sample text with custom OpenType font");

                // Apply the custom font to the cell style if the font was loaded
                Style style = worksheet.Cells["A1"].GetStyle();
                if (File.Exists(fontPath))
                {
                    style.Font.Name = "MyCustomFont"; // Name defined inside the .otf file
                }
                style.Font.Size = 14;
                worksheet.Cells["A1"].SetStyle(style);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    DefaultFont = File.Exists(fontPath) ? "MyCustomFont" : "Arial",
                    CheckWorkbookDefaultFont = true,
                    FontEncoding = PdfFontEncoding.Identity,
                    EmbedStandardWindowsFonts = true
                };

                // Save the workbook to PDF
                const string outputPath = "output.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}