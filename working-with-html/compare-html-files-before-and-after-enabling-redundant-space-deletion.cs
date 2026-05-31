using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlSpaceComparison
{
    class Program
    {
        static void Main()
        {
            // Prepare a sample HTML string containing redundant spaces.
            string htmlContent = "<p>   This    text   has   redundant   spaces   </p>";

            // Write the HTML to a temporary file for loading.
            string inputPath = "sample_input.html";
            File.WriteAllText(inputPath, htmlContent);

            // -----------------------------------------------------------------
            // 1. Load the HTML without deleting redundant spaces (default behavior).
            // -----------------------------------------------------------------
            // No HtmlLoadOptions are supplied, so DeleteRedundantSpaces remains false.
            Workbook wbDefault = new Workbook(inputPath);

            // Save the workbook back to HTML to observe the original spacing.
            string outputDefaultPath = "output_default.html";
            wbDefault.Save(outputDefaultPath, new HtmlSaveOptions());

            // -----------------------------------------------------------------
            // 2. Load the same HTML with DeleteRedundantSpaces enabled.
            // -----------------------------------------------------------------
            HtmlLoadOptions loadOpts = new HtmlLoadOptions();
            loadOpts.DeleteRedundantSpaces = true; // Enable redundant space removal.
            Workbook wbTrimmed = new Workbook(inputPath, loadOpts);

            // Save the trimmed workbook to HTML.
            string outputTrimmedPath = "output_trimmed.html";
            wbTrimmed.Save(outputTrimmedPath, new HtmlSaveOptions());

            // -----------------------------------------------------------------
            // 3. Compare the two generated HTML files.
            // -----------------------------------------------------------------
            string defaultHtml = File.ReadAllText(outputDefaultPath);
            string trimmedHtml = File.ReadAllText(outputTrimmedPath);

            Console.WriteLine("=== Comparison Result ===");
            Console.WriteLine($"Length before deletion : {defaultHtml.Length}");
            Console.WriteLine($"Length after deletion  : {trimmedHtml.Length}");
            Console.WriteLine();

            if (defaultHtml == trimmedHtml)
            {
                Console.WriteLine("The HTML files are identical.");
            }
            else
            {
                Console.WriteLine("The HTML files differ.");
                // Simple visual diff: show the cell text extracted from the workbook.
                string cellTextDefault = wbDefault.Worksheets[0].Cells["A1"].StringValue;
                string cellTextTrimmed = wbTrimmed.Worksheets[0].Cells["A1"].StringValue;

                Console.WriteLine($"Cell A1 text (default) : \"{cellTextDefault}\"");
                Console.WriteLine($"Cell A1 text (trimmed) : \"{cellTextTrimmed}\"");
            }

            // Clean up temporary files (optional).
            // File.Delete(inputPath);
            // File.Delete(outputDefaultPath);
            // File.Delete(outputTrimmedPath);
        }
    }
}