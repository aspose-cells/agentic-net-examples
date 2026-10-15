// Title: Convert English number words to integers and write them to an Excel worksheet using Aspose.Cells in C#
// AI Prompts: Generate C# code that creates a new workbook with Aspose.Cells, writes a list of English number words in column A, converts each word to its integer value, and writes the integers to column B. | Show how to build a case‑insensitive Dictionary<string,int> for number‑word mapping and use PutValue to insert both the original text and the numeric result into an .xlsx file with Aspose.Cells. | Provide a complete example that saves the populated workbook as "TextualToNumeric.xlsx" after performing the word‑to‑number conversion.
// Common Searches: aspnet convert number words to integers when writing to Excel with Aspose.Cells | c# Aspose.Cells example mapping English number words to numeric values in a worksheet | how to use a dictionary to translate word numbers to ints and save as .xlsx using Aspose.Cells | populate Excel column with text and corresponding numeric values using Aspose.Cells C#
// Tags: Aspose.Cells number‑word mapping C# | worksheet population using dictionary Aspose.Cells | PutValue for numeric data Excel .xlsx | case‑insensitive text‑to‑int lookup C# | export transformed data with Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;

// // Creates a new workbook, uses a case‑insensitive dictionary to translate English number words to integers, writes the words to column A and the integers to column B, and saves the file as TextualToNumeric.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample textual number representations
        string[] textualNumbers = { "Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten" };

        // Mapping from textual representation to numeric value
        Dictionary<string, int> textToNumber = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            { "Zero", 0 },
            { "One", 1 },
            { "Two", 2 },
            { "Three", 3 },
            { "Four", 4 },
            { "Five", 5 },
            { "Six", 6 },
            { "Seven", 7 },
            { "Eight", 8 },
            { "Nine", 9 },
            { "Ten", 10 }
        };

        // Populate the worksheet and convert textual numbers to numeric types
        for (int i = 0; i < textualNumbers.Length; i++)
        {
            // Column A: original textual representation
            sheet.Cells[i, 0].PutValue(textualNumbers[i]);

            // Convert to numeric using the dictionary; default to 0 if not found
            int numericValue = textToNumber.TryGetValue(textualNumbers[i], out int val) ? val : 0;

            // Column B: numeric value
            sheet.Cells[i, 1].PutValue(numericValue);
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("TextualToNumeric.xlsx");
    }
}
