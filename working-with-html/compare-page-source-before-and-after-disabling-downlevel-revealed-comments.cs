using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlCommentComparison
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a sample comment for comparison.";

            // -----------------------------------------------------------------
            // Save HTML with default behavior (DisableDownlevelRevealedComments = false)
            // -----------------------------------------------------------------
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions();
            defaultOptions.DisableDownlevelRevealedComments = false; // explicit, default is false
            string defaultHtmlPath = "output_default.html";
            workbook.Save(defaultHtmlPath, defaultOptions);

            // -----------------------------------------------------------------
            // Save HTML with downlevel-revealed comments disabled (true)
            // -----------------------------------------------------------------
            HtmlSaveOptions disabledOptions = new HtmlSaveOptions();
            disabledOptions.DisableDownlevelRevealedComments = true;
            string disabledHtmlPath = "output_disabled.html";
            workbook.Save(disabledHtmlPath, disabledOptions);

            // -----------------------------------------------------------------
            // Load the generated HTML files as plain text
            // -----------------------------------------------------------------
            string defaultHtml = File.ReadAllText(defaultHtmlPath);
            string disabledHtml = File.ReadAllText(disabledHtmlPath);

            // -----------------------------------------------------------------
            // Simple comparison: check if the contents are identical
            // -----------------------------------------------------------------
            bool areEqual = string.Equals(defaultHtml, disabledHtml, StringComparison.Ordinal);

            Console.WriteLine("Comparison result:");
            Console.WriteLine($"Are the two HTML files identical? {(areEqual ? "Yes" : "No")}");

            // If they differ, output a short snippet showing the first difference
            if (!areEqual)
            {
                int diffIndex = FindFirstDifference(defaultHtml, disabledHtml);
                int snippetLength = 200; // characters to display around the difference
                int start = Math.Max(diffIndex - snippetLength / 2, 0);
                int length = Math.Min(snippetLength, Math.Min(defaultHtml.Length, disabledHtml.Length) - start);

                Console.WriteLine("\nSnippet around first difference (default | disabled):");
                Console.WriteLine(defaultHtml.Substring(start, length));
                Console.WriteLine("---");
                Console.WriteLine(disabledHtml.Substring(start, length));
            }
        }

        // Helper method to locate the first differing character index between two strings
        private static int FindFirstDifference(string s1, string s2)
        {
            int minLength = Math.Min(s1.Length, s2.Length);
            for (int i = 0; i < minLength; i++)
            {
                if (s1[i] != s2[i])
                    return i;
            }
            return minLength; // either strings are equal up to minLength or one is a prefix of the other
        }
    }
}