// Title: Count distinct merged cell regions in an Aspose.Cells worksheet using C#
// AI Prompts: Write C# code that scans a worksheet with Aspose.Cells, uses GetMergedRange for each cell, and returns the number of unique merged areas. | Create a method that collects the first‑row/first‑column coordinates of every merged range into a HashSet and prints the total merged region count. | Modify the example to handle an arbitrary number of merged ranges and output the distinct merged region total without double‑counting.
// Common Searches: C# Aspose.Cells how to get total number of merged regions in a worksheet | count unique merged cells using GetMergedRange Aspose.Cells | determine distinct merged areas in Excel file with Aspose.Cells C# | enumerate used cells and identify merged ranges Aspose.Cells example | calculate merged region count programmatically Aspose.Cells
// Tags: count merged regions Aspose.Cells | GetMergedRange enumeration worksheet | unique merged area identifiers HashSet | Aspose.Cells merged range detection C# | calculate distinct merged cells Excel workbook

using Aspose.Cells;
using System;
using System.Collections.Generic;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The sample creates a workbook, merges several ranges, iterates over the used cells, uses GetMergedRange to detect merged areas, stores each area's top‑left coordinates in a HashSet to ensure uniqueness, prints the total number of distinct merged regions, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample merged regions
            cells.Merge(0, 0, 2, 2); // A1:B2
            cells.Merge(3, 1, 3, 3); // B4:D6
            cells.Merge(6, 0, 1, 5); // A7:F7

            // HashSet to store unique identifiers of merged regions (top‑left cell coordinates)
            HashSet<string> mergedRegions = new HashSet<string>();

            // Enumerate cells within the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    // GetMergedRange returns null if the cell is not part of a merged area
                    AsposeRange mergedRange = cell.GetMergedRange();
                    if (mergedRange != null)
                    {
                        // Use the first row and column of the merged area as a unique key
                        string key = $"{mergedRange.FirstRow}_{mergedRange.FirstColumn}";
                        mergedRegions.Add(key);
                    }
                }
            }

            // Output the total number of merged regions found
            Console.WriteLine($"Total merged regions: {mergedRegions.Count}");

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            string outputPath = "MergedRegionsCount.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
