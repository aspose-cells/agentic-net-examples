// Title: C# program to compare two HTML files after stripping downlevel revealed conditional comments
// Common Searches: how to remove IE downlevel revealed conditional comments from HTML using C# | C# compare two HTML files while ignoring conditional comments | regex pattern to strip downlevel revealed comments in C# | line by line diff of HTML sources after cleaning conditional comments in C# | programmatically compare cleaned page source files in .NET
// Tags: downlevel revealed comment removal C# | HTML source diff ignoring IE conditional comments | C# regular expression for conditional comments | line‑by‑line HTML comparison C# | cleaned HTML file comparison .NET

using System;
using System.IO;
using System.Text.RegularExpressions;

// The utility reads two HTML files, removes downlevel revealed conditional comments with a case‑insensitive single‑line regex, then compares the cleaned content. It reports whether the sources are identical or outputs line‑by‑line differences, enabling reliable HTML comparison after IE‑specific comments are stripped.
class PageSourceComparer
{
    // Regular expression to match downlevel revealed conditional comments:
    // <!--[if ...]> ... <![endif]-->
    private static readonly Regex DownlevelRevealedCommentRegex = new Regex(
        @"<!--\s*\[if[^\]]*\]>.*?<!\s*\[endif\]\s*-->",
        RegexOptions.Singleline | RegexOptions.IgnoreCase);

    static void Main(string[] args)
    {
        // Expect two file paths: original page source and modified page source
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: PageSourceComparer <original.html> <modified.html>");
            return;
        }

        string originalPath = args[0];
        string modifiedPath = args[1];

        if (!File.Exists(originalPath) || !File.Exists(modifiedPath))
        {
            Console.WriteLine("One or both files do not exist.");
            return;
        }

        // Load page sources
        string originalSource = File.ReadAllText(originalPath);
        string modifiedSource = File.ReadAllText(modifiedPath);

        // Disable (remove) downlevel revealed comments from both sources
        string cleanedOriginal = DisableDownlevelRevealedComments(originalSource);
        string cleanedModified = DisableDownlevelRevealedComments(modifiedSource);

        // Compare the cleaned sources
        if (cleanedOriginal == cleanedModified)
        {
            Console.WriteLine("The page sources are identical after disabling downlevel revealed comments.");
        }
        else
        {
            Console.WriteLine("Differences detected after disabling downlevel revealed comments:");
            ShowDifferences(cleanedOriginal, cleanedModified);
        }
    }

    // Removes downlevel revealed conditional comments from the HTML source
    private static string DisableDownlevelRevealedComments(string html)
    {
        return DownlevelRevealedCommentRegex.Replace(html, string.Empty);
    }

    // Simple line‑by‑line diff output
    private static void ShowDifferences(string original, string modified)
    {
        var originalLines = original.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        var modifiedLines = modified.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        int maxLines = Math.Max(originalLines.Length, modifiedLines.Length);

        for (int i = 0; i < maxLines; i++)
        {
            string oLine = i < originalLines.Length ? originalLines[i] : string.Empty;
            string mLine = i < modifiedLines.Length ? modifiedLines[i] : string.Empty;

            if (!string.Equals(oLine, mLine, StringComparison.Ordinal))
            {
                Console.WriteLine($"Line {i + 1}:");
                Console.WriteLine($"  Original: {oLine}");
                Console.WriteLine($"  Modified: {mLine}");
            }
        }
    }
}
