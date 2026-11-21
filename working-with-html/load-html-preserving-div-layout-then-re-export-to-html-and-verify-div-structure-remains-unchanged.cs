// Title: Load an HTML file with DIV layout using Aspose.Cells for .NET, re‑export to HTML and verify the DIV structure stays unchanged
// Common Searches: asp.net load html preserving div layout with Aspose.Cells | c# compare original and exported html to ensure div structure is unchanged using Aspose.Cells | how to use HtmlLoadOptions and HtmlSaveOptions to keep div layout when saving workbook as html | verify div‑based html export from an Aspose.Cells workbook in C# | aspocells memory stream save html without converting divs to tables
// Tags: HtmlLoadOptions preserve div layout | HtmlSaveOptions ExportHtmlAsTable false | Aspose.Cells load html workbook | Aspose.Cells export html using divs | C# normalize html strings for comparison | memory stream workbook save as html

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an HTML file that uses DIVs for layout into an Aspose.Cells Workbook with HtmlLoadOptions, saving the workbook back to HTML using HtmlSaveOptions (default DIV layout), writing the result through a MemoryStream, normalizing both the original and saved HTML to ignore whitespace, and comparing the strings to confirm that the DIV structure remains unchanged.
class HtmlDivPreserveDemo
{
    static void Main()
    {
        // Paths for input and output HTML files
        string inputHtmlPath = "input.html";
        string outputHtmlPath = "output.html";

        // Verify that the input HTML file exists
        if (!File.Exists(inputHtmlPath))
        {
            Console.WriteLine($"Input file not found: {inputHtmlPath}");
            return;
        }

        try
        {
            // Read original HTML content for later verification
            string originalHtml = File.ReadAllText(inputHtmlPath);

            // Load the HTML file into a Workbook
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

            // Prepare HTML save options.
            // By default Aspose.Cells exports using DIV layout when ExportHtmlAsTable is false.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // If the library version supports ExportHtmlAsTable, you can uncomment the line below:
            // saveOptions.ExportHtmlAsTable = false; // Use DIVs for layout

            // Save the workbook back to HTML using a memory stream
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, saveOptions);
                ms.Position = 0;

                // Write the saved HTML to the output file
                using (StreamWriter writer = new StreamWriter(outputHtmlPath, false))
                {
                    writer.Write(new StreamReader(ms).ReadToEnd());
                }

                // Retrieve the saved HTML content for verification
                ms.Position = 0;
                string savedHtml = new StreamReader(ms).ReadToEnd();

                // Simple verification: compare the original and saved HTML strings
                bool isDivStructurePreserved = string.Equals(
                    NormalizeHtml(originalHtml),
                    NormalizeHtml(savedHtml),
                    StringComparison.Ordinal);

                Console.WriteLine(isDivStructurePreserved
                    ? "DIV structure preserved after re-export."
                    : "DIV structure changed after re-export.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to normalize HTML strings (remove whitespace differences)
    static string NormalizeHtml(string html)
    {
        // Remove line breaks and excess spaces for a basic comparison
        return string.Join(" ", html
            .Replace("\r", " ")
            .Replace("\n", " ")
            .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
    }
}
