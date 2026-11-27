// Title: Replace <DATE> placeholders in every TextBox shape of an Excel workbook with the current system date using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace <DATE> tag in all TextBox objects across worksheets | batch update TextBox text in Excel file to current date with Aspose.Cells .NET | C# code to iterate through worksheets and modify TextBox content using Aspose.Cells | how to substitute custom placeholder in Excel shape text using Aspose.Cells API | replace multiple <DATE> placeholders in Excel TextBoxes programmatically
// Tags: Aspose.Cells TextBox content substitution | C# batch modify Excel TextBox shapes | update Excel shape placeholders with system date | worksheet iteration for TextBox text changes | save modified workbook after TextBox updates

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ReplaceDateInTextBoxes
{
    // The example loads an Excel workbook with Aspose.Cells, walks through each worksheet and its TextBox collection, replaces every "<DATE>" placeholder in the TextBox text with the current system date formatted as yyyy‑MM‑dd, tracks the number of replacements, and saves the updated workbook to a new file.
    class Program
    {
        /// <param name="args">Command‑line arguments.</param>
        static void Main(string[] args)
        {
            Console.WriteLine("=== Aspose.Cells - Replace <DATE> tags in all TextBoxes ===");

            // Validate arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Error: No input file specified.");
                Console.WriteLine("Usage: ReplaceDateInTextBoxes.exe <input.xlsx> [output.xlsx]");
                return;
            }

            string inputPath = args[0];
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            string outputPath;
            if (args.Length >= 2 && !string.IsNullOrWhiteSpace(args[1]))
            {
                outputPath = args[1];
            }
            else
            {
                string directory = Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory();
                string filenameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                outputPath = Path.Combine(directory, $"{filenameWithoutExt}_Updated.xlsx");
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook: \"{inputPath}\"");

                // Prepare the replacement string once
                string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
                Console.WriteLine($"Current system date (yyyy-MM-dd): {currentDate}");

                int totalReplacements = 0;

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    TextBoxCollection textBoxes = sheet.TextBoxes;
                    if (textBoxes == null || textBoxes.Count == 0)
                        continue; // No text boxes on this sheet

                    // Iterate through each TextBox in the collection
                    for (int i = 0; i < textBoxes.Count; i++)
                    {
                        TextBox tb = textBoxes[i];
                        if (tb == null)
                            continue;

                        string originalText = tb.Text;
                        if (string.IsNullOrEmpty(originalText))
                            continue;

                        if (originalText.Contains("<DATE>"))
                        {
                            string updatedText = originalText.Replace("<DATE>", currentDate);
                            tb.Text = updatedText;

                            int replacementsInThisBox = CountOccurrences(originalText, "<DATE>");
                            totalReplacements += replacementsInThisBox;

                            Console.WriteLine($"Worksheet \"{sheet.Name}\", TextBox index {i}: Replaced {replacementsInThisBox} occurrence(s).");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: \"{outputPath}\"");
                Console.WriteLine($"Total <DATE> tag replacements performed: {totalReplacements}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <summary>
        /// Helper method to count how many times a substring appears in a string.
        /// </summary>
        /// <param name="source">The source string.</param>
        /// <param name="substring">The substring to count.</param>
        /// <returns>Number of occurrences.</returns>
        private static int CountOccurrences(string source, string substring)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(substring))
                return 0;

            int count = 0;
            int index = 0;
            while ((index = source.IndexOf(substring, index, StringComparison.Ordinal)) != -1)
            {
                count++;
                index += substring.Length;
            }
            return count;
        }
    }
}
