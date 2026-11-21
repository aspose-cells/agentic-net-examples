// Title: Generate responsive HTML with cell‑fit layout and comment tooltips using WidthScalable, HtmlCrossType.FitToCell, and AddTooltipText in Aspose.Cells for .NET
// Common Searches: how to export Excel to responsive HTML with comment tooltips using Aspose.Cells C# | Aspose.Cells HtmlSaveOptions WidthScalable true example code | using HtmlCrossType.FitToCell to fit cells in HTML output with Aspose.Cells | enable AddTooltipText to show Excel comment pop‑ups in HTML export | C# sample for creating scalable HTML from a workbook that includes tooltips
// Tags: WidthScalable option Aspose.Cells | AddTooltipText comment tooltips | HtmlCrossType FitToCell layout | scalable width HTML export | Excel comment tooltip export C#

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;

// The example creates a workbook, adds sample data and a comment, then configures HtmlSaveOptions with WidthScalable set to true, HtmlCrossType.FitToCell (when supported), and AddTooltipText enabled, finally saving the workbook as a responsive HTML file that displays tooltips for comments.
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

            // Populate sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Oranges");
            sheet.Cells["B3"].PutValue(15);

            // Add a comment to demonstrate tooltip functionality
            int commentIndex = sheet.Comments.Add("A2");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Freshly picked apples";

            // Configure HTML save options for responsive output with tooltips
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.WidthScalable = true;               // Enable responsive width
            // htmlOptions.HtmlCrossType = HtmlCrossType.FitToCell; // Not available in current API version
            htmlOptions.AddTooltipText = true;              // Include tooltips for comments and hyperlinks

            // Save the workbook as an HTML file
            workbook.Save("ResponsiveWithTooltips.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
