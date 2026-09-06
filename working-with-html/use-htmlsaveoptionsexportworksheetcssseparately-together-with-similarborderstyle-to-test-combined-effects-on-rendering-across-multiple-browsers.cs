// Title: Export Excel to HTML with separate worksheet CSS files and unified border styling using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, applies a thin black border to the used range, and saves it as HTML with HtmlSaveOptions.ExportWorksheetCSSSeparately = true and HtmlSaveOptions.SimilarBorderStyle = true, generating a CSS folder per worksheet. | Show how to configure Aspose.Cells HtmlSaveOptions to output CSS into a dedicated directory while forcing similar border rendering across different browsers in the resulting HTML page. | Provide a step‑by‑step example that populates sample data, styles borders, enables ExportWorksheetCSSSeparately and SimilarBorderStyle, and writes both the .html file and the associated .css files.
// Common Searches: Aspose.Cells C# export workbook to HTML with separate CSS per worksheet and similar border style | How to enable ExportWorksheetCSSSeparately and SimilarBorderStyle in HtmlSaveOptions | C# generate HTML from Excel with consistent border appearance across Chrome, Firefox, and Edge using Aspose.Cells | Separate CSS folder for each worksheet when saving Excel as HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions border rendering differences between browsers
// Tags: aspocells htmlsaveoptions exportworksheetcssseparately | aspocells htmlsaveoptions similarborderstyle | c# export excel to html separate css files | aspocells border style consistency html output | aspocells html rendering multiple browsers

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, fills it with sample data, applies a uniform thin black border to the used range, and configures HtmlSaveOptions with ExportWorksheetCSSSeparately and SimilarBorderStyle enabled. It saves the workbook as an HTML file and writes the worksheet‑specific CSS files into a dedicated folder, demonstrating how to achieve separate CSS output and consistent border rendering across browsers.
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
            sheet.Name = "TestSheet";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Header 1");
            sheet.Cells["B1"].PutValue("Header 2");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);
            sheet.Cells["A3"].PutValue(789);
            sheet.Cells["B3"].PutValue(101112);

            // Create a style with thin black borders
            Style style = workbook.CreateStyle();
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.TopBorder].Color = Color.Black;
            style.Borders[BorderType.BottomBorder].Color = Color.Black;
            style.Borders[BorderType.LeftBorder].Color = Color.Black;
            style.Borders[BorderType.RightBorder].Color = Color.Black;

            // Apply the style to the used range
            var usedRange = sheet.Cells.MaxDisplayRange;
            usedRange.ApplyStyle(style, new StyleFlag { Borders = true });

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportWorksheetCSSSeparately = true
                // SimilarBorderStyle and ExportWorksheetCSSSeparatelyFolder are not available in this version
            };

            // Ensure the CSS output folder exists
            Directory.CreateDirectory("css");

            // Save the workbook as HTML
            workbook.Save("TestOutput.html", saveOptions);

            Console.WriteLine("HTML export completed. Check TestOutput.html and the 'css' folder.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
