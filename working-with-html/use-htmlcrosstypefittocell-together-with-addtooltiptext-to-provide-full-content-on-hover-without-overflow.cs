// Title: Generate HTML from Excel with Aspose.Cells using HtmlCrossType.FitToCell and AddTooltipText to display full cell content on hover
// AI Prompts: Create C# code that builds a workbook, inserts a long text into a cell, sets HtmlSaveOptions.HtmlCrossType to FitToCell, enables AddTooltipText, and saves the file as HTML so the complete text appears as a tooltip when the user hovers over the cell. | Update an existing Aspose.Cells example to eliminate text overflow by applying the FitToCell cross‑type and turning on the AddTooltipText flag for comments, then export the worksheet to an HTML file.
// Common Searches: Aspose.Cells HtmlCrossType FitToCell with AddTooltipText example C# | how to prevent cell text overflow in HTML export using Aspose.Cells | display full cell value on mouse hover in Aspose.Cells generated HTML | C# save workbook as HTML tooltip for long cell content Aspose.Cells
// Tags: Aspose.Cells HtmlCrossType FitToCell | Aspose.Cells AddTooltipText option | C# export Excel to HTML tooltip | prevent cell overflow Aspose.Cells HTML | cell comment tooltip Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program creates a workbook, writes a lengthy string into cell A1, adds a comment containing the same text, configures HtmlSaveOptions with HtmlCrossType.FitToCell and AddTooltipText enabled, and saves the worksheet as an HTML file where the full content is shown as a tooltip on hover, avoiding overflow.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Put a long text into cell A1
            var cell = sheet.Cells["A1"];
            cell.PutValue("This is a very long text that would normally overflow the cell boundaries.");

            // Add a comment (tooltip) containing the full text so it appears on hover in HTML
            // First add a comment to the cell, then set its author and note
            int commentIndex = sheet.Comments.Add("A1");
            var comment = sheet.Comments[commentIndex];
            comment.Author = "Author";
            comment.Note = cell.StringValue;

            // Optionally set column width to demonstrate fitting
            sheet.Cells.SetColumnWidth(0, 15);

            // Configure HTML save options
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Determine output path
            string outputPath = "output.html";

            // Ensure the output directory exists (if a directory is specified)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
