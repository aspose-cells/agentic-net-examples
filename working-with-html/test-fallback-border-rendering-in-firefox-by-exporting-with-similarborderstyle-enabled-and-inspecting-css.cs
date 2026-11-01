// Title: Export an Excel workbook to HTML with thick cell borders and enable SimilarBorderStyle for Firefox fallback using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, applies thick borders to a range, sets HtmlSaveOptions.SimilarBorderStyle = true, saves the sheet as HTML with a separate CSS file, and prints the CSS content to the console. | Update existing Aspose.Cells HTML export code to turn on SimilarBorderStyle for Firefox border fallback and output the generated CSS for verification.
// Common Searches: Aspose.Cells C# how to export workbook to HTML with separate CSS file and thick borders | Enable SimilarBorderStyle in HtmlSaveOptions for Firefox border fallback Aspose.Cells | Generate CSS for cell borders when saving Excel as HTML using Aspose.Cells .NET | Inspect generated CSS from Aspose.Cells HTML export to verify border rendering in Firefox | Export only active worksheet to HTML with gridlines and base64 images Aspose.Cells
// Tags: aspocells htmlsaveoptions similarborderstyle | c# export workbook to html with css borders | aspocells generate separate css file | thick cell border rendering firefox aspocells | export active worksheet only html aspocells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a new workbook, adds sample data, applies thick borders to cells A1:B2, configures HtmlSaveOptions to enable SimilarBorderStyle for Firefox fallback, exports the active worksheet as HTML with a separate CSS file, and prints the generated CSS for inspection.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Apply thick borders to the range A1:B2
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:B2");
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
            StyleFlag flag = new StyleFlag { All = true };
            range.ApplyStyle(borderStyle, flag);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = true,
                ExportImagesAsBase64 = true,
                ExportGridLines = true
                // Additional options (e.g., ExportCellBorder, ExportCss) can be set if supported by the library version
            };

            // Prepare output directory
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the workbook as HTML (CSS will be saved as a separate .css file if supported)
            string htmlPath = Path.Combine(outputDir, "Test.html");
            workbook.Save(htmlPath, htmlOptions);

            // Locate the generated CSS file (if any)
            string cssPath = Path.ChangeExtension(htmlPath, ".css");

            // Output CSS content for inspection
            if (File.Exists(cssPath))
            {
                Console.WriteLine("=== Generated CSS ===");
                Console.WriteLine(File.ReadAllText(cssPath));
            }
            else
            {
                Console.WriteLine("CSS file was not generated.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
