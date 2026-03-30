using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsRegexDemo
{
    class Program
    {
        static void Main()
        {
            // Define the regular expression pattern
            string pattern = @"^[pb][aeiou]";

            // Instantiate a compiled Regex object for efficient reuse
            // RegexOptions.Compiled improves performance for repeated matches
            // RegexOptions.CultureInvariant ensures consistent behavior across cultures
            Regex regex = new Regex(pattern, RegexOptions.Compiled | RegexOptions.CultureInvariant);

            // Sample data to test the regex
            string[] samples = { "apple", "banana", "pear", "plum", "orange" };

            // Perform matching operations using the pre‑compiled regex
            foreach (string text in samples)
            {
                bool isMatch = regex.IsMatch(text);
                Console.WriteLine($"{text}: {(isMatch ? "matches" : "does not match")} the pattern");
            }
        }
    }
}