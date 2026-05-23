using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCustomLookup
{
    class Program
    {
        // Searches all worksheets for the specified value and returns a list of
        // 1‑cell ranges that contain matching cells.
        static List<AsposeRange> GetMatchingRanges(Workbook workbook, string searchValue)
        {
            var matchedRanges = new List<AsposeRange>();
            try
            {
                // Configure find options: search in cell values, exact match.
                var findOptions = new FindOptions
                {
                    LookInType = LookInType.Values,
                    LookAtType = LookAtType.EntireContent
                };

                // Iterate through each worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cell previous = null;
                    Cell found = sheet.Cells.Find(searchValue, previous, findOptions);

                    while (found != null)
                    {
                        // Create a 1‑cell range for the found cell.
                        AsposeRange cellRange = sheet.Cells.CreateRange(found.Row, found.Column, 1, 1);
                        matchedRanges.Add(cellRange);

                        // Continue searching after the current cell.
                        previous = found;
                        found = sheet.Cells.Find(searchValue, previous, findOptions);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during GetMatchingRanges: {ex.Message}");
            }

            return matchedRanges;
        }

        static void Main()
        {
            try
            {
                // Load template workbook safely.
                string templatePath = "Template.xlsx";
                Workbook wb = File.Exists(templatePath) ? new Workbook(templatePath) : new Workbook();

                // Add a second worksheet.
                int sheet2Index = wb.Worksheets.Add();
                Worksheet sheet1 = wb.Worksheets[0];
                Worksheet sheet2 = wb.Worksheets[sheet2Index];

                // Populate sample data.
                sheet1.Cells["A1"].PutValue("Alpha");
                sheet1.Cells["B2"].PutValue("Target");
                sheet1.Cells["C3"].PutValue("Gamma");

                sheet2.Cells["A1"].PutValue("Target");
                sheet2.Cells["B2"].PutValue("Delta");
                sheet2.Cells["C3"].PutValue("Target");

                // Find all cells containing "Target".
                List<AsposeRange> resultRanges = GetMatchingRanges(wb, "Target");

                if (resultRanges.Count > 0)
                {
                    // Define highlight style.
                    Style highlight = wb.CreateStyle();
                    highlight.ForegroundColor = Color.Yellow;
                    highlight.Pattern = BackgroundType.Solid;

                    // Apply style to each matching range.
                    foreach (var range in resultRanges)
                    {
                        range.SetStyle(highlight);
                    }
                }

                // Save the workbook.
                string outputPath = "CustomLookupResult.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}