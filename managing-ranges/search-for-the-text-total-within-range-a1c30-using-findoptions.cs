// Title: How to find the text "Total" in the A1:C30 range using Aspose.Cells FindOptions in C#
// Common Searches: Aspose.Cells FindOptions search for a specific string within a defined range in C# | C# code to locate the word 'Total' in cells A1 to C30 using Aspose.Cells | example of iterating over a range with FindOptions to find text in an Excel worksheet | search Excel cells for exact text with case‑insensitive FindOptions in Aspose.Cells .NET | retrieve cell addresses of matching text using Aspose.Cells FindOptions in C#
// Tags: Aspose.Cells FindOptions string search | search Excel range A1:C30 C# | case-insensitive cell value lookup Aspose.Cells | retrieve matching CellArea objects Aspose.Cells | save workbook after search Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, fills cells A1:C3 with sample data, configures a FindOptions object for case‑insensitive value search, manually scans rows 0‑29 and columns 0‑2 (A1:C30), compares each string cell to "Total", collects matching cells as CellArea objects, prints their row/column addresses and values, and saves the workbook as SearchResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data (optional)
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Quantity");
            cells["C1"].PutValue("Total");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["C2"].PutValue(20);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(5);
            cells["C3"].PutValue(15);
            // Additional rows can be added as needed

            // Configure FindOptions (essential properties)
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,          // Search in cell values
                CaseSensitive = false                    // Case‑insensitive
                // LookAtType is omitted because we perform a whole‑text comparison manually
            };

            // Define search range: rows 0‑29, columns 0‑2 (A1:C30)
            int startRow = 0, startColumn = 0, endRow = 29, endColumn = 2;

            // Manual search because FindAll may not be available in some versions
            List<CellArea> foundList = new List<CellArea>();
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startColumn; col <= endColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell != null && cell.Type == CellValueType.IsString)
                    {
                        string cellText = cell.StringValue;
                        // Apply FindOptions criteria
                        bool match = findOptions.CaseSensitive
                            ? string.Equals(cellText, "Total", StringComparison.Ordinal)
                            : string.Equals(cellText, "Total", StringComparison.OrdinalIgnoreCase);

                        if (match)
                        {
                            // Create CellArea for the single cell
                            CellArea area = new CellArea
                            {
                                StartRow = row,
                                StartColumn = col,
                                EndRow = row,
                                EndColumn = col
                            };
                            foundList.Add(area);
                        }
                    }
                }
            }

            CellArea[] foundAreas = foundList.Count > 0 ? foundList.ToArray() : null;

            // Process the search results
            if (foundAreas != null && foundAreas.Length > 0)
            {
                foreach (CellArea area in foundAreas)
                {
                    // Output the address of each found cell (1‑based indexing for readability)
                    Console.WriteLine($"Found at: Row {area.StartRow + 1}, Column {area.StartColumn + 1}");
                    // Retrieve and display the cell's value
                    Cell foundCell = cells[area.StartRow, area.StartColumn];
                    Console.WriteLine($"Value: {foundCell.StringValue}");
                }
            }
            else
            {
                Console.WriteLine("Text 'Total' not found in the specified range.");
            }

            // Save the workbook (optional)
            string outputPath = "SearchResult.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
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
