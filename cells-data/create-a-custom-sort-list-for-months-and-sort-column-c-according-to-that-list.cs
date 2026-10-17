// Title: How to sort an Excel column of month names using a custom month order with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to sort the values in column C of a worksheet according to a predefined chronological month list. | Show how to build a case‑insensitive dictionary that maps month names to their order index and apply it to sort a string array before writing the sorted data back to the sheet. | Adapt the example to sort multiple columns by the same custom month order within a single workbook using Aspose.Cells.
// Common Searches: aspnet sort Excel column by custom month sequence using Aspose.Cells | c# Aspose.Cells custom sort order for month names in a worksheet | how to implement a predefined month list for sorting column data with Aspose.Cells .NET | using dictionary lookup to sort month strings in Aspose.Cells workbook
// Tags: Aspose.Cells month list sorting | C# dictionary driven Excel sort | custom column order sorting .NET | chronological month order in workbook

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example creates a new workbook, writes an unsorted list of month names to column C, defines the correct chronological month order, builds a case‑insensitive dictionary to map each month to its index, sorts the month array using this custom order, overwrites column C with the sorted values, optionally sets an AutoFilter range, and saves the file as SortedMonths.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample unsorted month names in column C (index 2)
            string[] unsortedMonths = {
                "March", "January", "December", "July", "May",
                "October", "February", "August", "June",
                "November", "April", "September"
            };

            // Write unsorted months to the sheet (rows 0‑11, column C)
            for (int i = 0; i < unsortedMonths.Length; i++)
            {
                sheet.Cells[i, 2].PutValue(unsortedMonths[i]); // Column C
            }

            // Define the chronological order for months
            string[] monthOrder = new string[]
            {
                "January", "February", "March", "April", "May", "June",
                "July", "August", "September", "October", "November", "December"
            };

            // Build a lookup dictionary for quick index retrieval
            var monthIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < monthOrder.Length; i++)
            {
                monthIndex[monthOrder[i]] = i;
            }

            // Sort the month names using the custom order
            Array.Sort(unsortedMonths, (a, b) =>
            {
                int idxA = monthIndex.ContainsKey(a) ? monthIndex[a] : int.MaxValue;
                int idxB = monthIndex.ContainsKey(b) ? monthIndex[b] : int.MaxValue;
                return idxA.CompareTo(idxB);
            });

            // Write the sorted months back to the sheet (overwrites previous values)
            for (int i = 0; i < unsortedMonths.Length; i++)
            {
                sheet.Cells[i, 2].PutValue(unsortedMonths[i]);
            }

            // Optionally set an AutoFilter range (not required for sorting)
            sheet.AutoFilter.Range = "A1:C12";

            // Save the workbook
            string outputPath = "SortedMonths.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
