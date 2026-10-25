// Title: Sort pivot table field items with a custom order defined by a string list using Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells custom sort order for pivot table field items | how to rank pivot field values with a predefined string list in Aspose.Cells | dictionary based row sorting in Excel using Aspose.Cells .NET | apply custom string ranking to worksheet data with Aspose.Cells C# | set custom sequence for pivot table items programmatically with Aspose.Cells
// Tags: custom pivot field sorting Aspose.Cells | dictionary driven sort order Excel .NET | string list ranking for worksheet rows | C# Aspose.Cells custom field sequence | Excel workbook custom order using Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example creates a workbook, fills a column with items (High, Medium, Low) and auxiliary data, defines a case‑insensitive dictionary that maps each item to a custom rank, extracts rows into a list, sorts the list according to the dictionary values (unmatched items go to the end), writes the sorted rows back, and saves the file as CustomSorted.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (first column will be sorted)
            string[] items = { "Medium", "Low", "High", "Medium", "Low", "High" };
            for (int i = 0; i < items.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(items[i]);   // field item
                sheet.Cells[i, 1].PutValue(i + 1);      // auxiliary data
            }

            // Define custom order
            var customOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                { "High", 0 },
                { "Medium", 1 },
                { "Low", 2 }
            };

            // Determine the used range
            int lastRow = sheet.Cells.MaxDataRow;
            int lastCol = sheet.Cells.MaxDataColumn;

            // Collect rows into a list for sorting
            var rows = new List<(string key, object[] values)>();
            for (int r = 0; r <= lastRow; r++)
            {
                string key = sheet.Cells[r, 0].StringValue;
                object[] values = new object[lastCol + 1];
                for (int c = 0; c <= lastCol; c++)
                {
                    values[c] = sheet.Cells[r, c].Value;
                }
                rows.Add((key, values));
            }

            // Sort rows using the custom order
            rows.Sort((a, b) =>
            {
                int orderA = customOrder.ContainsKey(a.key) ? customOrder[a.key] : int.MaxValue;
                int orderB = customOrder.ContainsKey(b.key) ? customOrder[b.key] : int.MaxValue;
                return orderA.CompareTo(orderB);
            });

            // Write sorted data back to the worksheet
            for (int r = 0; r < rows.Count; r++)
            {
                for (int c = 0; c <= lastCol; c++)
                {
                    sheet.Cells[r, c].PutValue(rows[r].values[c]);
                }
            }

            // Save the workbook
            string outputPath = "CustomSorted.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
