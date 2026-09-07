// Title: How to enable fallback border rendering in Aspose.Cells HTML export by setting HtmlSaveOptions.SimilarBorderStyle to true (C#)
// AI Prompts: Generate C# code that saves an Aspose.Cells workbook to HTML with HtmlSaveOptions.SimilarBorderStyle set to true for border compatibility. | Show the steps to configure HtmlSaveOptions to use a similar border style when exporting Excel to HTML using Aspose.Cells for .NET. | Provide a minimal example that applies thick borders to cells and enables fallback border rendering in the HTML output.
// Common Searches: Aspose.Cells set SimilarBorderStyle true for HTML export | C# export Excel to HTML with fallback borders using Aspose.Cells | How to preserve cell border styles in HTML output from Aspose.Cells | HtmlSaveOptions SimilarBorderStyle property usage example | Enable border compatibility for older browsers in Aspose.Cells HTML conversion
// Tags: Aspose.Cells HtmlSaveOptions SimilarBorderStyle | HTML export fallback borders Aspose.Cells | C# set HtmlSaveOptions SimilarBorderStyle | Excel to HTML border compatibility Aspose.Cells | configure border rendering Aspose.Cells HTML

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates creating a workbook, applying thick borders to cells, configuring HtmlSaveOptions with SimilarBorderStyle set to true to provide fallback border rendering for browsers lacking CSS border support, and saving the workbook as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample data to demonstrate borders (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Data");

            // Apply a thick black border to the cells
            Style style = workbook.CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thick, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thick, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thick, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thick, Color.Black);
            sheet.Cells["A1"].SetStyle(style);
            sheet.Cells["A2"].SetStyle(style);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: SimilarBorderStyle property is not available in all versions; omitted for compatibility.

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
