// Title: C# – Set Excel workbook Subject built‑in property to the most frequent word in the first worksheet using Aspose.Cells
// Common Searches: Aspose.Cells C# set Subject built‑in property from worksheet text | How to update Excel document properties with the most common word in a sheet | Extract most frequent word from Excel cells and assign to Subject using .NET | Programmatically count word frequency in an Excel worksheet with Aspose.Cells | Automatically generate Excel metadata based on sheet content in C#
// Tags: set Subject built‑in property Aspose.Cells | word frequency analysis Excel worksheet C# | update Excel document properties programmatically | extract most common word from cells Aspose.Cells | auto generate Excel metadata .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// The example loads an Excel file (or creates a new workbook if the file is missing or protected), scans all string cells in the first worksheet, counts words longer than three characters, selects the most frequent word (defaulting to "Document"), assigns that word to the workbook's Subject built‑in document property, and saves the workbook, with error handling for loading, property setting, and saving.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        Workbook workbook = null;

        // Load workbook safely, handling missing file or password protection
        try
        {
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook if the input file does not exist
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }
        }
        catch (CellsException)
        {
            // If the file is password‑protected or another CellsException occurs, fall back to a new workbook
            workbook = new Workbook();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Ensure workbook is not null (should never happen, but guard anyway)
        if (workbook == null)
        {
            Console.WriteLine("Failed to create or load a workbook.");
            return;
        }

        // Analyze the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        var wordCount = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (Cell cell in cells)
        {
            if (cell.Type == CellValueType.IsString)
            {
                string[] words = cell.StringValue.Split(
                    new char[] { ' ', ',', '.', ';', ':', '-', '\t', '\r', '\n' },
                    StringSplitOptions.RemoveEmptyEntries);

                foreach (string w in words)
                {
                    string word = w.Trim();
                    if (word.Length < 4) continue; // ignore very short words

                    if (wordCount.ContainsKey(word))
                        wordCount[word]++;
                    else
                        wordCount[word] = 1;
                }
            }
        }

        // Determine the most frequent word as the main topic
        string mainTopic = "Document";
        if (wordCount.Count > 0)
        {
            mainTopic = wordCount.OrderByDescending(kv => kv.Value).First().Key;
        }

        // Update the Subject built‑in property
        try
        {
            workbook.BuiltInDocumentProperties["Subject"].Value = mainTopic;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error setting document property: {ex.Message}");
        }

        // Save the workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
