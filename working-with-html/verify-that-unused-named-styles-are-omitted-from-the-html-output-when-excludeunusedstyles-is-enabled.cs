// Title: How to omit unused named styles from HTML output using Aspose.Cells HtmlSaveOptions in C#
// Common Searches: C# Aspose.Cells how to remove unused styles when saving as HTML | example of using HtmlSaveOptions.ExcludeUnusedStyles in Aspose.Cells | check if Aspose.Cells HTML export excludes unused named styles | Aspose.Cells HTML export omit unused CSS styles C# | validate that only used styles appear in generated HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions exclude unused styles | C# export worksheet to HTML without unused styles | verify CSS style removal in Aspose.Cells HTML output | HTML conversion with style pruning using Aspose.Cells | unused named style omission in Aspose.Cells HTML export

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The sample creates a workbook with one applied style and one unused style, saves the worksheet to HTML twice—first with ExcludeUnusedStyles disabled and then enabled—and reads the resulting files to confirm that the used style's CSS is present while the unused style's CSS is omitted when the option is turned on.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for the generated HTML files
            string htmlWithUnusedPath = "HtmlWithUnusedStyles.html";
            string htmlWithoutUnusedPath = "HtmlWithoutUnusedStyles.html";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Create a style that will be used
            // -------------------------------------------------
            Style usedStyle = workbook.CreateStyle();
            usedStyle.Font.Color = Color.Red;
            usedStyle.Font.Size = 12;

            // Apply the used style to a cell
            sheet.Cells["A1"].PutValue("This cell uses the used style");
            sheet.Cells["A1"].SetStyle(usedStyle);

            // -------------------------------------------------
            // 2. Create a style that will NOT be used
            // -------------------------------------------------
            Style unusedStyle = workbook.CreateStyle();
            unusedStyle.Font.Color = Color.Blue;
            unusedStyle.Font.Size = 14;
            // Note: we do NOT apply this style to any cell

            // -------------------------------------------------
            // 3. Save HTML with ExcludeUnusedStyles = false (default)
            // -------------------------------------------------
            HtmlSaveOptions includeAllOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true,
                ExcludeUnusedStyles = false // keep all styles
            };
            workbook.Save(htmlWithUnusedPath, includeAllOptions);
            Console.WriteLine($"HTML saved with all styles to '{htmlWithUnusedPath}'.");

            // -------------------------------------------------
            // 4. Save HTML with ExcludeUnusedStyles = true
            // -------------------------------------------------
            HtmlSaveOptions excludeUnusedOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true,
                ExcludeUnusedStyles = true // omit unused styles
            };
            workbook.Save(htmlWithoutUnusedPath, excludeUnusedOptions);
            Console.WriteLine($"HTML saved excluding unused styles to '{htmlWithoutUnusedPath}'.");

            // -------------------------------------------------
            // 5. Verify the presence/absence of the style definitions in the HTML output
            // -------------------------------------------------
            VerifyStylePresence(htmlWithUnusedPath, "color: red", "color: blue");
            VerifyStylePresence(htmlWithoutUnusedPath, "color: red", "color: blue");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    /// <param name="htmlPath">Path to the HTML file.</param>
    /// <param name="usedStyleSnippet">CSS snippet that should be present (used style).</param>
    /// <param name="unusedStyleSnippet">CSS snippet that should be absent when ExcludeUnusedStyles is true.</param>
    private static void VerifyStylePresence(string htmlPath, string usedStyleSnippet, string unusedStyleSnippet)
    {
        try
        {
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"ERROR: File '{htmlPath}' does not exist.");
                return;
            }

            string htmlContent = File.ReadAllText(htmlPath);

            bool containsUsed = htmlContent.IndexOf(usedStyleSnippet, StringComparison.OrdinalIgnoreCase) >= 0;
            bool containsUnused = htmlContent.IndexOf(unusedStyleSnippet, StringComparison.OrdinalIgnoreCase) >= 0;

            Console.WriteLine($"\nVerification for '{htmlPath}':");
            Console.WriteLine($"  Contains used style snippet \"{usedStyleSnippet}\": {(containsUsed ? "YES" : "NO")}");
            Console.WriteLine($"  Contains unused style snippet \"{unusedStyleSnippet}\": {(containsUnused ? "YES" : "NO")}");

            if (containsUsed && !containsUnused)
            {
                Console.WriteLine("  Result: SUCCESS – unused styles are omitted as expected.");
            }
            else if (containsUsed && containsUnused)
            {
                Console.WriteLine("  Result: FAILURE – unused styles are still present.");
            }
            else if (!containsUsed)
            {
                Console.WriteLine("  Result: FAILURE – the used style is missing (unexpected).");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Verification error for '{htmlPath}': {ex.Message}");
        }
    }
}
