// Title: Export an Aspose.Cells workbook to HTML with gridlines enabled and a custom TableCssId for border styling (C#)
// AI Prompts: Write C# code that creates a workbook, applies thin borders to the used range, and saves it as HTML with gridlines enabled and a custom CSS ID for the table using Aspose.Cells. | Show how to configure Aspose.Cells HTML export options to retain cell borders by turning on gridlines and assigning a table identifier. | Demonstrate exporting an Excel worksheet to HTML while preserving border styling, using Aspose.Cells settings for gridlines and table CSS prefix.
// Common Searches: how to enable gridlines in Aspose.Cells HTML export C# | Aspose.Cells HtmlSaveOptions custom table id example | preserve Excel cell borders when converting to HTML with Aspose.Cells | export workbook to HTML with custom table identifier using Aspose.Cells | C# Aspose.Cells retain borders in HTML output
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | Aspose.Cells HtmlSaveOptions TableCssId | Aspose.Cells preserve cell borders HTML | Aspose.Cells apply borders used range | Aspose.Cells export workbook to HTML C#

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // Creates a workbook, adds sample data, applies thin black borders to the used range, configures HtmlSaveOptions with gridlines enabled and TableCssId set to "myTable", and saves the workbook as an HTML file, preserving cell borders.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data
                sheet.Cells["A1"].PutValue("Header 1");
                sheet.Cells["B1"].PutValue("Header 2");
                sheet.Cells["A2"].PutValue("Data 1");
                sheet.Cells["B2"].PutValue("Data 2");

                // Apply a simple style to demonstrate borders
                Style style = workbook.CreateStyle();
                style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                style.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                style.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                style.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;

                // Apply the style to the used range
                var usedRange = sheet.Cells.MaxDisplayRange; // Aspose.Cells.Range inferred
                usedRange.ApplyStyle(style, new StyleFlag() { All = true });

                // Configure HTML save options to enable gridlines and set a CSS ID prefix for the table
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Enable gridlines so that cell borders are retained in the exported HTML
                    ExportGridLines = true,

                    // Prefix the generated table with a CSS ID to allow custom styling of borders
                    TableCssId = "myTable"
                };

                // Save the workbook as HTML using the configured options
                workbook.Save("ExportedWithGridlines.html", htmlOptions);
                Console.WriteLine("Workbook exported successfully to ExportedWithGridlines.html");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
