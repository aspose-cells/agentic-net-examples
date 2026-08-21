// Title: C# custom Aspose.Cells function to find a value across all worksheets and return its Range
// Description: Demonstrates how to create a static FindFirstMatchRange method that iterates through every worksheet in an Aspose.Cells Workbook, uses FindOptions (LookInType.Values, LookAtType.Contains) to locate the first cell containing a specified text, returns a one‑cell Range, highlights the cell, and saves the workbook. Includes sample data on two sheets (Products and Sales).
// Keywords: Aspose.Cells | C# | .NET | custom lookup function | search across worksheets | FindOptions | Range object | highlight cell | Excel automation | Workbook | find cell value
// Common Searches: Aspose.Cells find value in any worksheet | C# return Range for first matching cell across sheets | custom function to search all sheets Aspose.Cells | highlight found cell Aspose.Cells C# | search multiple worksheets and get cell range
// Developer Intent: Locate the first occurrence of a given text in any worksheet of a workbook and obtain it as an Aspose.Cells Range for further processing.
// Use Cases: Identify and highlight a product name that may appear on several sheets before generating a report. | Validate data consistency by detecting duplicate entries across multiple worksheets. | Extract the cell range of a matching value to feed into formulas, conditional formatting, or chart data sources.
// AI Prompts: Write a C# method using Aspose.Cells that searches for a string in every worksheet and returns the first matching cell as a one‑cell Range. | Show how to call the custom lookup method, apply a yellow background style to the returned range, and save the workbook to a given file path. | Explain how to adapt FindFirstMatchRange for exact matches, case‑insensitive searches, or whole‑sheet scans.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCustomLookup
{
    // Demonstrates how to create a static FindFirstMatchRange method that iterates through every worksheet in an Aspose.Cells Workbook, uses FindOptions (LookInType.Values, LookAtType.Contains) to locate the first cell containing a specified text, returns a one‑cell Range, highlights the cell, and saves the workbook. Includes sample data on two sheets (Products and Sales).
    class Program
    {
        // Custom function that searches for a value across all worksheets
        // and returns the first matching cell as an AsposeRange object.
        static AsposeRange? FindFirstMatchRange(Workbook workbook, string searchValue)
        {
            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Configure find options: search in cell values and allow partial matches
                FindOptions options = new FindOptions
                {
                    LookInType = LookInType.Values,
                    LookAtType = LookAtType.Contains
                };

                // Perform the search starting from the beginning (previousCell = null)
                Cell foundCell = sheet.Cells.Find(searchValue, null, options);

                // If a matching cell is found, create and return a one‑cell range
                if (foundCell != null)
                {
                    return sheet.Cells.CreateRange(foundCell.Row, foundCell.Column, 1, 1);
                }
            }

            // No match found in any worksheet
            return null;
        }

        static void Main()
        {
            try
            {
                // Create a new workbook and add sample data on two worksheets
                Workbook wb = new Workbook();

                // Worksheet 1
                Worksheet ws1 = wb.Worksheets[0];
                ws1.Name = "Products";
                ws1.Cells["A1"].PutValue("Item");
                ws1.Cells["A2"].PutValue("Apple");
                ws1.Cells["A3"].PutValue("Banana");

                // Worksheet 2 (Add returns a Worksheet when a name is supplied)
                Worksheet ws2 = wb.Worksheets.Add("Sales");
                ws2.Cells["B1"].PutValue("Product");
                ws2.Cells["B2"].PutValue("Orange");
                ws2.Cells["B3"].PutValue("Apple"); // Duplicate value to test cross‑sheet search

                // Use the custom lookup function to find "Apple"
                string lookupValue = "Apple";
                AsposeRange? resultRange = FindFirstMatchRange(wb, lookupValue);

                if (resultRange != null)
                {
                    // Highlight the found range for visual confirmation
                    Style highlight = wb.CreateStyle();
                    highlight.ForegroundColor = Color.Yellow;
                    highlight.Pattern = BackgroundType.Solid;
                    resultRange.SetStyle(highlight);

                    Console.WriteLine($"Found '{lookupValue}' in worksheet '{resultRange.Worksheet.Name}' at {resultRange.FirstRow},{resultRange.FirstColumn}");
                }
                else
                {
                    Console.WriteLine($"Value '{lookupValue}' not found in any worksheet.");
                }

                // Save the workbook
                string outputPath = "CustomLookupResult.xlsx";

                // Ensure the directory exists before saving
                string? directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
