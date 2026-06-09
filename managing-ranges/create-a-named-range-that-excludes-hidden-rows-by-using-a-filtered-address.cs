using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class NamedRangeExcludingHiddenRows
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (header + 6 rows)
                cells["A1"].PutValue("Category");
                cells["A2"].PutValue("Apple");
                cells["A3"].PutValue("Banana");
                cells["A4"].PutValue("Apple");
                cells["A5"].PutValue("Cherry");
                cells["A6"].PutValue("Apple");
                cells["A7"].PutValue("Date");

                // Apply an AutoFilter to the header row covering column A
                sheet.AutoFilter.Range = "A1:A7";

                // Filter to show only rows where the value is "Apple"
                sheet.AutoFilter.AddFilter(0, "Apple");
                // Refresh the filter; hideRows = true hides the filtered rows
                sheet.AutoFilter.Refresh(true);

                // Get the indexes of hidden rows (0‑based)
                int[] hiddenRows = sheet.AutoFilter.Refresh(false);

                // Determine the visible rows range (excluding hidden rows)
                int startRow = 1; // data starts at row index 1 (A2)
                int endRow = sheet.Cells.MaxDataRow; // last row with data
                var visibleRanges = new List<(int start, int end)>();
                int currentStart = -1;

                for (int row = startRow; row <= endRow; row++)
                {
                    bool isHidden = false;
                    if (hiddenRows != null)
                    {
                        foreach (int h in hiddenRows)
                        {
                            if (h == row)
                            {
                                isHidden = true;
                                break;
                            }
                        }
                    }

                    if (!isHidden)
                    {
                        if (currentStart == -1)
                            currentStart = row; // start a new visible block
                    }
                    else
                    {
                        if (currentStart != -1)
                        {
                            visibleRanges.Add((currentStart, row - 1));
                            currentStart = -1;
                        }
                    }
                }

                // Add the last block if it ends at the last row
                if (currentStart != -1)
                    visibleRanges.Add((currentStart, endRow));

                // Build the address string for the named range
                StringBuilder addressBuilder = new StringBuilder();
                for (int i = 0; i < visibleRanges.Count; i++)
                {
                    var range = visibleRanges[i];
                    // Convert zero‑based row indexes to Excel row numbers (add 1)
                    int firstRow = range.start + 1;
                    int lastRow = range.end + 1;
                    addressBuilder.Append($"{sheet.Name}!A{firstRow}:A{lastRow}");
                    if (i < visibleRanges.Count - 1)
                        addressBuilder.Append(",");
                }

                string filteredAddress = addressBuilder.ToString();

                // Create a named range that refers only to the visible rows
                int nameIndex = workbook.Worksheets.Names.Add("VisibleApple");
                Name visibleName = workbook.Worksheets.Names[nameIndex];
                visibleName.RefersTo = "=" + filteredAddress;

                // Demonstrate that the named range works in a formula
                cells["B1"].Formula = "=SUM(VisibleApple)";
                workbook.CalculateFormula();

                Console.WriteLine($"Named range 'VisibleApple' refers to: {visibleName.RefersTo}");
                Console.WriteLine($"Sum of visible rows (should count only Apple rows): {cells["B1"].Value}");

                // Save the workbook
                workbook.Save("NamedRangeExcludingHiddenRows.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            NamedRangeExcludingHiddenRows.Run();
        }
    }
}