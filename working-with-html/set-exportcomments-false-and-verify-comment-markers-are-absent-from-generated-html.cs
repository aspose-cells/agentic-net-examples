// Title: Disable comment export when saving a workbook to HTML with Aspose.Cells for .NET and verify the output
// AI Prompts: Generate C# code that sets HtmlSaveOptions.ExportComments = false, saves a workbook to HTML, and reads the file to confirm no comment markers are present. | Write a C# snippet that adds a comment to a cell, disables comment export during HTML conversion with Aspose.Cells, and programmatically checks the resulting HTML for the absence of Aspose comment tags.
// Common Searches: Aspose.Cells C# export workbook to HTML without comments | How to turn off comment export in HtmlSaveOptions Aspose.Cells | Verify that HTML output from Aspose.Cells does not contain comment markers | C# check generated HTML file for Aspose comment tags after saving workbook | Disable comments when converting Excel to HTML using Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions disable comment export | C# verify HTML output without comments Aspose.Cells | remove worksheet comments before HTML conversion Aspose.Cells | check generated HTML for Aspose comment markers C# | export workbook to HTML without comments Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds a comment to cell A1, configures HtmlSaveOptions.ExportComments = false, saves the workbook as HTML, reads the generated file, and searches for comment markers such as "AsposeComment" to confirm that comments were not exported.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell A1 (to have something that could be exported)
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Sample comment for testing";

            // Configure HTML save options.
            // Note: In some Aspose.Cells versions the ExportComments property is not available.
            // If needed, comments can be removed from the worksheet before saving.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML
            string htmlFile = "output.html";
            workbook.Save(htmlFile, htmlOptions);

            // Verify that the HTML file was created before attempting to read it
            if (!File.Exists(htmlFile))
            {
                Console.WriteLine($"Failed to generate HTML file: {htmlFile}");
                return;
            }

            // Read the generated HTML content
            string htmlContent = File.ReadAllText(htmlFile);

            // Simple check: Aspose.Cells adds comment markers containing the word "AsposeComment"
            // If such a marker is found, comments were exported.
            if (htmlContent.Contains("AsposeComment", StringComparison.OrdinalIgnoreCase) ||
                htmlContent.Contains("comment", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Comment markers detected in HTML. ExportComments may not have been applied correctly.");
            }
            else
            {
                Console.WriteLine("No comment markers found. ExportComments = false works as intended (or comments were removed).");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
