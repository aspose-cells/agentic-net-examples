// Title: Loading HTML with Aspose.Cells while preserving DIV layout, exporting to HTML with Base64 images, and verifying DIV structure in C#
// AI Prompts: Load an HTML file into a Workbook using HtmlLoadOptions, then save it with HtmlSaveOptions where ExportImagesAsBase64 is set to true. | Read the original and generated HTML files, count opening <div> tags case‑insensitively, and output whether the counts match to confirm layout preservation. | Add robust error handling that checks for the input file's existence and logs any exceptions occurring during load or save.
// Common Searches: Aspose.Cells keep div elements unchanged after HTML load and save in C# | C# compare number of <div> tags before and after Aspose.Cells HTML export | How to save HTML with images encoded as Base64 using Aspose.Cells .NET | Validate HTML layout consistency after round‑trip conversion using Aspose.Cells | Example of HtmlLoadOptions and HtmlSaveOptions for HTML round‑trip in C#
// Tags: HtmlLoadOptions preserve DIV hierarchy | HtmlSaveOptions ExportImagesAsBase64 | Aspose.Cells HTML round‑trip verification | C# count opening DIV tags | Workbook load and save HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an HTML file into an Aspose.Cells Workbook with HtmlLoadOptions, saves it back to HTML using HtmlSaveOptions that embed images as Base64, then reads both files and compares the count of opening <div> tags to ensure the DIV layout remains unchanged.
class Program
{
    static void Main()
    {
        // Paths to the source and destination HTML files
        string inputPath = "input.html";
        string outputPath = "output.html";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the HTML file. HtmlLoadOptions automatically treats the file as HTML.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook back to HTML while embedding images as Base64.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportImagesAsBase64 = true
                // ExportHtmlAsMhtml defaults to false (standard HTML), so no need to set it explicitly.
            };
            workbook.Save(outputPath, saveOptions);

            // ----- Verification of DIV structure -----
            // Read both original and regenerated HTML content.
            string originalHtml = File.ReadAllText(inputPath);
            string regeneratedHtml = File.ReadAllText(outputPath);

            // Simple verification: compare the count of opening <div> tags.
            int originalDivCount = CountTagOccurrences(originalHtml, "div");
            int regeneratedDivCount = CountTagOccurrences(regeneratedHtml, "div");

            Console.WriteLine($"Original DIV count: {originalDivCount}");
            Console.WriteLine($"Regenerated DIV count: {regeneratedDivCount}");
            Console.WriteLine(originalDivCount == regeneratedDivCount
                ? "DIV structure preserved."
                : "DIV structure changed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to count opening tags (e.g., <div> or <div ...>).
    static int CountTagOccurrences(string html, string tagName)
    {
        if (string.IsNullOrEmpty(html) || string.IsNullOrEmpty(tagName))
            return 0;

        // Count occurrences of "<tag" (covers <tag> and <tag ...>)
        string pattern = $"<{tagName}";
        int count = 0;
        int index = 0;

        while ((index = html.IndexOf(pattern, index, StringComparison.OrdinalIgnoreCase)) != -1)
        {
            count++;
            index += pattern.Length;
        }

        return count;
    }
}
