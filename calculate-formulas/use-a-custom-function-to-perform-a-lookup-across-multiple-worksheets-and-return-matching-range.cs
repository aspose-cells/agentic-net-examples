// Title: Custom lookup across all worksheets with Aspose.Cells for .NET (C#) – find and highlight matching cells
// Description: Demonstrates a reusable FindAcrossWorksheets method that scans every worksheet in an Aspose.Cells Workbook, uses FindOptions (LookInType.Values, LookAtType.Contains) to locate cells containing a specified string, creates a 1 × 1 Range for each hit, returns a List<Range>, and shows how to apply a yellow style, output the addresses, and save the result.
// Keywords: Aspose.Cells find across worksheets | C# custom lookup function | highlight matching cells Aspose.Cells | FindOptions LookInType.Values | search multiple sheets .NET | create range from found cell | Excel data extraction C# | Aspose.Cells Find method example
// Common Searches: Aspose.Cells search all worksheets for a value | C# highlight cells that match a keyword in Excel | How to return cell ranges from Find in Aspose.Cells | Custom lookup function for multiple sheets Aspose.Cells | Find and style matching cells across worksheets .NET
// Developer Intent: Create a reusable C# function that searches every worksheet in a workbook for a given text, returns the matching cell ranges, and enables further actions such as styling or reporting.
// Use Cases: Locate and highlight every occurrence of a product name across product and sales sheets. | Generate a list of cell addresses that contain a specific keyword for audit or reporting. | Apply a uniform formatting (e.g., yellow background) to all cells matching a lookup term before exporting the workbook.
// AI Prompts: Write a C# method using Aspose.Cells that searches for a string in all worksheets and returns a List<Range> of the matches. | Show how to apply a yellow background style to each range returned by a custom lookup function in Aspose.Cells. | Explain how to modify the FindAcrossWorksheets function to perform case‑insensitive searches and restrict results to a particular column.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomLookup
{
    // Demonstrates a reusable FindAcrossWorksheets method that scans every worksheet in an Aspose.Cells Workbook, uses FindOptions (LookInType.Values, LookAtType.Contains) to locate cells containing a specified string, creates a 1 × 1 Range for each hit, returns a List<Range>, and shows how to apply a yellow style, output the addresses, and save the result.
    class Program
    {
        // Searches for a value across all worksheets and returns a list of ranges
        // that represent the matching cells.
        static List<Aspose.Cells.Range> FindAcrossWorksheets(Workbook workbook, string searchValue)
        {
            var options = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains
            };

            var result = new List<Aspose.Cells.Range>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cell previous = null;
                Cell found = sheet.Cells.Find(searchValue, previous, options);

                while (found != null)
                {
                    // Create a 1x1 range for the found cell.
                    var cellRange = sheet.Cells.CreateRange(found.Row, found.Column, 1, 1);
                    result.Add(cellRange);

                    // Continue searching from the current cell.
                    previous = found;
                    found = sheet.Cells.Find(searchValue, previous, options);
                }
            }

            return result; // Empty list if no matches were found.
        }

        static void Main()
        {
            try
            {
                // Create a new workbook and populate data.
                var wb = new Workbook();

                var ws1 = wb.Worksheets[0];
                ws1.Name = "Products";
                ws1.Cells["A1"].PutValue("Item");
                ws1.Cells["A2"].PutValue("Apple");
                ws1.Cells["A3"].PutValue("Banana");
                ws1.Cells["A4"].PutValue("Cherry");

                int sheetIndex = wb.Worksheets.Add();
                var ws2 = wb.Worksheets[sheetIndex];
                ws2.Name = "Sales";
                ws2.Cells["B1"].PutValue("Product");
                ws2.Cells["B2"].PutValue("Apple");
                ws2.Cells["B3"].PutValue("Durian");
                ws2.Cells["B4"].PutValue("Banana");

                // Perform the custom lookup for the term "Apple".
                var matches = FindAcrossWorksheets(wb, "Apple");

                if (matches.Count > 0)
                {
                    // Highlight all matching cells.
                    var highlight = wb.CreateStyle();
                    highlight.ForegroundColor = Color.Yellow;
                    highlight.Pattern = BackgroundType.Solid;

                    foreach (var range in matches)
                    {
                        range.SetStyle(highlight);
                    }

                    Console.WriteLine("Matches found at:");
                    foreach (var range in matches)
                    {
                        Console.WriteLine($"{range.Worksheet.Name}!{range.RefersTo}");
                    }
                }
                else
                {
                    Console.WriteLine("No matches found.");
                }

                // Save the workbook.
                string outputPath = "LookupResult.xlsx";

                // Ensure the directory exists before saving.
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
