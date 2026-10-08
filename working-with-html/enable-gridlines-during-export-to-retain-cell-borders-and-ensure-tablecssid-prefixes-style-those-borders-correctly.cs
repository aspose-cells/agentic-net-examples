// Title: Enable gridlines and specify TableCssId when exporting an Aspose.Cells workbook to HTML to keep cell borders (C#)
// AI Prompts: Write C# code that creates a workbook, applies thin black borders to a range, and saves it as HTML with ExportGridLines enabled and a custom TableCssId. | Show how to configure HtmlSaveOptions in Aspose.Cells to retain cell borders by exporting gridlines and prefixing CSS selectors with a table ID. | Demonstrate applying a border style to cells and exporting the worksheet to HTML while preserving the borders using the Aspose.Cells C# API.
// Common Searches: Aspose.Cells how to export workbook to HTML with gridlines and custom table id | C# preserve cell borders in HTML output using Aspose.Cells HtmlSaveOptions | Enable ExportGridLines and set TableCssId in Aspose.Cells HTML export example | Apply thin borders to a range and export to HTML with Aspose.Cells C# | Why cell borders disappear when saving Aspose.Cells workbook as HTML
// Tags: export workbook to html with gridlines Aspose.Cells | set TableCssId HtmlSaveOptions Aspose.Cells | apply thin border style to range Aspose.Cells | preserve cell borders in html export C# | configure HtmlSaveOptions ExportGridLines true

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// Alias to avoid ambiguity with System.Range
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, adds sample data, applies thin black borders to cells A1:B2, configures HtmlSaveOptions to export gridlines and use a custom TableCssId, and saves the result as an HTML file.
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

            // Fill some sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data2");

            // Define a style with thin black borders
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
            borderStyle.Borders[BorderType.RightBorder].Color = Color.Black;
            borderStyle.Borders[BorderType.TopBorder].Color = Color.Black;
            borderStyle.Borders[BorderType.BottomBorder].Color = Color.Black;

            // Apply the border style to the used range (A1:B2)
            AsposeRange usedRange = sheet.Cells.CreateRange("A1:B2");
            usedRange.ApplyStyle(borderStyle, new StyleFlag { All = true });

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export gridlines so that cell borders appear in the exported HTML
                ExportGridLines = true,
                // Prefix CSS selectors with this table ID to style borders correctly
                TableCssId = "myTable"
            };

            // Ensure output directory exists
            string outputPath = "output.html";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the workbook to HTML using the configured options
            workbook.Save(outputPath, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
