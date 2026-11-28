// Title: C# Aspose.Cells: Convert an Excel (.xlsx) workbook to HTML with PresentationPreference.BestFit while preserving cell comments
// Common Searches: how to save an Excel workbook as HTML with best‑fit layout using Aspose.Cells in C# | Aspose.Cells export .xlsx to .html keeping cell comments | set PresentationPreference.BestFit for HTML conversion with Aspose.Cells C# | C# example converting Excel to HTML and exporting visible comments with Aspose.Cells | Aspose.Cells HtmlSaveOptions preserve worksheet comments during conversion
// Tags: Aspose.Cells HtmlSaveOptions best fit presentation | PresentationPreference.BestFit Aspose.Cells | convert .xlsx to .html preserving cell comments | C# Aspose.Cells HTML conversion example | export Excel comments to HTML using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample program checks for an input.xlsx file, creates a workbook with a visible comment if needed, loads the workbook, configures HtmlSaveOptions with PresentationPreference set to true for a best‑fit layout, and saves the workbook as output.html. Cell comments are automatically included in the generated HTML.
class Program
{
    static void Main()
    {
        // Paths for the input Excel file and the output HTML file
        string inputPath = "input.xlsx";
        string outputPath = "output.html";

        try
        {
            // Ensure a sample workbook exists if the input file is missing
            if (!File.Exists(inputPath))
            {
                CreateSampleWorkbookWithComment(inputPath);
                Console.WriteLine($"Sample workbook created at '{inputPath}'.");
            }

            // Load the workbook (file existence already verified)
            Workbook workbook = new Workbook(inputPath);
            Console.WriteLine($"Workbook '{inputPath}' loaded successfully.");

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Set to true for best‑fit presentation (more beautiful layout)
                PresentationPreference = true
                // Cell comments are exported by default when saving to HTML
            };

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully converted to HTML at '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during conversion:");
            Console.WriteLine(ex.Message);
        }
    }

    /// <param name="path">The file path where the workbook will be saved.</param>
    private static void CreateSampleWorkbookWithComment(string path)
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Put some data into a cell
            ws.Cells["A1"].PutValue("Hello Aspose!");

            // Add a comment to the same cell
            int commentIndex = ws.Comments.Add("A1");
            Comment comment = ws.Comments[commentIndex];
            comment.Note = "This is a sample comment.";
            comment.IsVisible = true; // Make the comment visible

            // Save the workbook
            wb.Save(path);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to create sample workbook:");
            Console.WriteLine(ex.Message);
            throw;
        }
    }
}
