// Title: Create a custom C# lookup function in Aspose.Cells to find and highlight matching cells across all worksheets
// AI Prompts: Write a C# method that uses Aspose.Cells FindOptions to search every worksheet in a workbook for a given string and returns a List<Cell> of all matches. | Modify the lookup method to accept a case‑insensitive flag and apply a user‑defined Style (e.g., background color) to each found cell. | Generate code that calls the custom lookup, highlights the results, and saves the workbook to a specified file path.
// Common Searches: aspnet cells find text in all sheets and highlight results | c# Aspose.Cells search across multiple worksheets for a value | how to apply background color to cells found by custom lookup in Aspose.Cells | retrieve list of cells containing a substring using Aspose.Cells FindOptions | save workbook after styling matched cells with Aspose.Cells C#
// Tags: Aspose.Cells custom lookup across worksheets | FindOptions Contains search Aspose.Cells | highlight matching cells Aspose.Cells style | collect List<Cell> from multi‑sheet search | save workbook after cell styling Aspose.Cells | case‑insensitive text search Aspose.Cells C#

using System;
using System.Collections.Generic;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsCustomLookup
{
    // The example defines a C# helper that iterates through every worksheet in an Aspose.Cells workbook, uses FindOptions with a Contains match to locate all cells containing a specified string, gathers those Cell objects into a List<Cell>, applies a light‑yellow background style to each, and finally saves the workbook as an Excel file.
    class Program
    {
        // Custom function that searches for a value across all worksheets
        // and returns a list of all matching cells.
        static List<Cell> GetMatchingCells(Workbook workbook, string searchValue)
        {
            var result = new List<Cell>();

            // Configure find options – search in cell values and allow partial matches.
            var findOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains
            };

            // Iterate through each worksheet in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                Cell previous = null;

                // Keep finding next occurrence until none is left.
                while (true)
                {
                    Cell found = cells.Find(searchValue, previous, findOptions);
                    if (found == null)
                        break;

                    result.Add(found);
                    previous = found;
                }
            }

            return result;
        }

        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle rule: create).
                var wb = new Workbook();

                // Populate first worksheet.
                Worksheet ws1 = wb.Worksheets[0];
                ws1.Name = "Products";
                ws1.Cells["A1"].PutValue("Apple");
                ws1.Cells["A2"].PutValue("Banana");
                ws1.Cells["A3"].PutValue("Cherry");

                // Add a second worksheet.
                int sheetIndex = wb.Worksheets.Add();
                Worksheet ws2 = wb.Worksheets[sheetIndex];
                ws2.Name = "Sales";
                ws2.Cells["B1"].PutValue("Apple");
                ws2.Cells["B2"].PutValue("Durian");
                ws2.Cells["B3"].PutValue("Banana");

                // Use the custom lookup to find all cells containing "Apple".
                List<Cell> matches = GetMatchingCells(wb, "Apple");

                // Highlight the found cells with a light yellow background.
                if (matches.Count > 0)
                {
                    Style highlight = wb.CreateStyle();
                    highlight.BackgroundColor = Color.LightYellow;

                    foreach (Cell cell in matches)
                    {
                        cell.SetStyle(highlight);
                    }
                }

                // Save the workbook (lifecycle rule: save).
                wb.Save("LookupResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
