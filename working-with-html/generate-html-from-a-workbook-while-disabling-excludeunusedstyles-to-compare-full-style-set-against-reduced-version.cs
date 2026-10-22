// Title: Export an Excel workbook to HTML twice—once preserving all styles and once with default unused‑style removal using Aspose.Cells for .NET
// AI Prompts: Write a C# program that loads a .xlsx file, sets HtmlSaveOptions.ExcludeUnusedStyles to false, and saves the workbook as a full‑style HTML document. | Show how to save the same workbook a second time with the default HtmlSaveOptions so that unused styles are omitted, producing a reduced‑style HTML file. | Add error handling for a missing input file and include code that prints the paths of the two generated HTML files for easy comparison.
// Common Searches: Aspose.Cells C# export Excel to HTML without removing unused styles | how to keep all cell styles when converting .xlsx to HTML using Aspose.Cells | compare full style HTML output vs reduced style output from Aspose.Cells | HtmlSaveOptions ExcludeUnusedStyles false example in .NET
// Tags: HtmlSaveOptions ExcludeUnusedStyles false | Aspose.Cells HTML export full style set | C# export Excel to HTML with all styles | full vs reduced HTML export Aspose.Cells | disable unused style removal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing Excel workbook and saves two HTML files: one with ExcludeUnusedStyles disabled to retain every style, and another using default options where unused styles are omitted, enabling a direct visual comparison of full‑style versus reduced‑style HTML output.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // -------------------------------------------------
            // Save HTML with all styles included (ExcludeUnusedStyles disabled)
            // -------------------------------------------------
            HtmlSaveOptions fullStyleOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExcludeUnusedStyles = false // Keep every style, even if unused
            };
            workbook.Save("output_full.html", fullStyleOptions);

            // -------------------------------------------------
            // Save HTML with unused styles excluded (default behavior)
            // -------------------------------------------------
            HtmlSaveOptions reducedStyleOptions = new HtmlSaveOptions(SaveFormat.Html);
            // ExcludeUnusedStyles defaults to true, so no need to set it explicitly
            workbook.Save("output_reduced.html", reducedStyleOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
