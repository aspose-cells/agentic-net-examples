// Title: Convert English word numbers to integers and write them to Excel cells using Aspose.Cells for .NET (C#)
// Common Searches: how to map English number words to integers and write them to an Excel file with Aspose.Cells in C# | Aspose.Cells C# example for converting textual numbers to numeric cell values | populate column A with numeric values from a word list while keeping the original text in column B using Aspose.Cells | C# code to convert word numbers to integers and save the workbook as .xlsx with Aspose.Cells | dictionary based conversion of number words to integers in an Aspose.Cells worksheet
// Tags: convert word numbers to integers Aspose.Cells | populate Excel cells with numeric values C# | dictionary mapping number words Aspose.Cells | PutValue method numeric assignment C# | save workbook as xlsx Aspose.Cells | case-insensitive number word lookup C#

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The sample creates a new workbook, defines a case‑insensitive dictionary that maps English word numbers (Zero‑Ten) to their integer equivalents, iterates over the list, writes each numeric value to column A and the original word to column B of the first worksheet, and saves the result as TextToNumber.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample textual numbers
        string[] textualNumbers = { "Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten" };

        // Mapping textual representation to numeric value
        Dictionary<string, int> numberMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            {"Zero",0},{"One",1},{"Two",2},{"Three",3},{"Four",4},{"Five",5},
            {"Six",6},{"Seven",7},{"Eight",8},{"Nine",9},{"Ten",10}
        };

        // Populate cells with numeric values converted from textual numbers
        for (int i = 0; i < textualNumbers.Length; i++)
        {
            string text = textualNumbers[i];
            int numericValue;
            if (numberMap.TryGetValue(text, out numericValue))
            {
                // Assign numeric value to column A (index 0)
                sheet.Cells[i, 0].PutValue(numericValue);
                // Keep original text in column B (index 1) for reference
                sheet.Cells[i, 1].PutValue(text);
            }
        }

        // Save the workbook
        workbook.Save("TextToNumber.xlsx");
    }
}
