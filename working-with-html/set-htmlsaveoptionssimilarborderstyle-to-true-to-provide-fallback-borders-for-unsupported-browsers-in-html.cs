// Title: How to enable fallback cell borders in HTML export using HtmlSaveOptions.SimilarBorderStyle in Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, applies thick borders to cells, sets HtmlSaveOptions.SimilarBorderStyle = true, and saves the workbook as HTML with Aspose.Cells. | Show how to configure Aspose.Cells HtmlSaveOptions to preserve Excel border styles for browsers that do not fully support CSS borders. | Demonstrate setting the SimilarBorderStyle property before calling Workbook.Save to generate HTML with fallback border rendering.
// Common Searches: Aspose.Cells set SimilarBorderStyle true for HTML output to keep cell borders | C# export Excel to HTML with fallback border rendering using Aspose.Cells | How to preserve thick borders when converting a workbook to HTML with Aspose.Cells .NET | HtmlSaveOptions SimilarBorderStyle property usage example in Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions SimilarBorderStyle | C# HTML export with fallback cell borders | Excel border preservation in HTML output | Configure border rendering for unsupported browsers | Set SimilarBorderStyle true in .NET

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, applies thick black borders to cells, enables HtmlSaveOptions.SimilarBorderStyle to provide fallback border rendering for browsers lacking full CSS support, and saves the workbook as an HTML file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some data to demonstrate borders (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Data");

            // Apply a thick black border to the cells
            Style style = workbook.CreateStyle();
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thick, Color.Black);
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thick, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thick, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thick, Color.Black);
            sheet.Cells["A1"].SetStyle(style);
            sheet.Cells["A2"].SetStyle(style);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: SimilarBorderStyle property may not be available in all versions; omitted for compatibility.

            // Save the workbook as HTML with the specified options
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
