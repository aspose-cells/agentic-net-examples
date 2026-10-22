// Title: C# console program to compare HTML page source before and after removing downlevel revealed conditional comments
// AI Prompts: Write a C# console application that loads two HTML files, strips downlevel revealed conditional comments using a regular expression, and outputs the differing lines with simple +/- markers. | Refactor the diff routine to produce a unified diff format (including context lines and @@ headers) instead of the basic line‑by‑line markers. | Enhance the tool to write each removed conditional comment block to a separate log file while still performing the source comparison.
// Common Searches: how to strip IE conditional comments from HTML in a C# application | C# code to compare two HTML documents after removing conditional comments | using Regex to delete downlevel revealed comments from HTML files | display differences between cleaned HTML sources in a .NET console app | example of line‑by‑line diff for HTML strings in C#
// Tags: downlevel revealed comment removal C# | HTML conditional comment regex .NET | compare cleaned HTML sources C# | line-by-line HTML diff console app | log stripped conditional blocks C#

using System;
using System.IO;
using System.Text.RegularExpressions;

// A C# console utility that reads two HTML files, removes downlevel revealed conditional comments via a regular expression, and compares the cleaned content line by line, reporting any differences.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Paths to the HTML files before and after disabling downlevel revealed comments
            string beforePath = "before.html";
            string afterPath = "after.html";

            // Verify that the input files exist
            if (!File.Exists(beforePath))
            {
                Console.WriteLine($"Input file not found: {beforePath}");
                return;
            }

            if (!File.Exists(afterPath))
            {
                Console.WriteLine($"Input file not found: {afterPath}");
                return;
            }

            // Load the page sources
            string beforeSource = File.ReadAllText(beforePath);
            string afterSource = File.ReadAllText(afterPath);

            // Remove downlevel revealed comments from both sources
            string cleanedBefore = RemoveDownlevelRevealedComments(beforeSource);
            string cleanedAfter = RemoveDownlevelRevealedComments(afterSource);

            // Compare the cleaned sources
            if (cleanedBefore == cleanedAfter)
            {
                Console.WriteLine("No differences detected after disabling downlevel revealed comments.");
            }
            else
            {
                Console.WriteLine("Differences detected:");
                ShowDifferences(cleanedBefore, cleanedAfter);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Removes downlevel revealed conditional comments from HTML content
    static string RemoveDownlevelRevealedComments(string html)
    {
        // Matches patterns like <!--[if ...]> ... <![endif]--> and removes them
        const string pattern = @"<!--\s*\[if[^\]]*\]>(.*?)<!\s*\[endif\]\s*-->";
        return Regex.Replace(html, pattern, string.Empty, RegexOptions.Singleline | RegexOptions.IgnoreCase);
    }

    // Simple line‑by‑line diff output
    static void ShowDifferences(string before, string after)
    {
        var beforeLines = before.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        var afterLines = after.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        int maxLines = Math.Max(beforeLines.Length, afterLines.Length);

        for (int i = 0; i < maxLines; i++)
        {
            string lineBefore = i < beforeLines.Length ? beforeLines[i] : string.Empty;
            string lineAfter = i < afterLines.Length ? afterLines[i] : string.Empty;

            if (lineBefore != lineAfter)
            {
                Console.WriteLine($"Line {i + 1}:\n- {lineBefore}\n+ {lineAfter}\n");
            }
        }
    }
}
