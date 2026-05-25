using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class UnicodeFontPdfDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("PDF generated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add Unicode text that may not be supported by the default system font
            sheet.Cells["A1"].PutValue("This text uses the default font");
            sheet.Cells["A2"].PutValue("如果默认字体支持中文，这将显示正确"); // Chinese characters

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set a font that supports the required Unicode characters (e.g., SimSun)
                DefaultFont = "SimSun",
                // Try to use the workbook's default font first
                CheckWorkbookDefaultFont = true,
                // Ensure font compatibility checking is enabled (default true)
                CheckFontCompatibility = true
            };

            // Define output file path
            string outputPath = "UnicodeFontOutput.pdf";

            // Ensure we can write to the output location
            try
            {
                workbook.Save(outputPath, pdfOptions);
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
                throw;
            }
        }
    }
}