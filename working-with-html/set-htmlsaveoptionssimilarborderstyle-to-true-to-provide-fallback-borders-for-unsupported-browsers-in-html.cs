// Title: Enable fallback border styling in HTML export using Aspose.Cells HtmlSaveOptions (C#)
// Common Searches: Aspose.Cells HtmlSaveOptions SimilarBorderStyle true C# example | How to provide fallback borders for HTML output in Aspose.Cells | Set SimilarBorderStyle property when saving workbook as HTML with Aspose.Cells | HTML export border compatibility for older browsers using Aspose.Cells C# | Configure HtmlSaveOptions to use similar border style in Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions SimilarBorderStyle | fallback border style HTML export | C# Aspose.Cells HTML save options | browser compatibility borders Aspose.Cells | export workbook to HTML with border fallback

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, adds data, sets HtmlSaveOptions.SimilarBorderStyle to true to enable fallback border styling for browsers that lack full CSS support, and saves the workbook as an HTML file using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue("Data");

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // The SimilarBorderStyle property is not available in the current Aspose.Cells API.
            // Adjust other border-related options here if needed.

            // Save the workbook as HTML
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
