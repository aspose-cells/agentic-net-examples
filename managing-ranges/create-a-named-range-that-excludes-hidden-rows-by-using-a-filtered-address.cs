// Title: C# – Create a Named Range Excluding Hidden Rows Using a Filtered Address with Aspose.Cells
// Description: Demonstrates how to build a workbook, apply an AutoFilter, collect the addresses of visible (non‑hidden) rows, combine them into a comma‑separated address, assign that address to a named range, use the name in a SUM formula, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells named range visible rows | exclude hidden rows Aspose.Cells | C# filtered address Excel | AutoFilter visible rows Aspose | Aspose.Cells sum visible data | .NET create named range | Excel named range hidden rows
// Common Searches: Aspose.Cells create named range from filtered rows | C# named range that skips hidden rows | How to reference only visible rows in a named range using Aspose.Cells | Build filtered address for named range in .NET | Sum visible rows after AutoFilter Aspose.Cells
// Developer Intent: Generate a named range that references only rows visible after an AutoFilter is applied.
// Use Cases: Automated reporting that needs to aggregate data from filtered rows only. | Dynamic dashboards where hidden rows must be excluded from calculations. | Exporting workbooks with custom named ranges for downstream analytics.
// AI Prompts: Write C# code with Aspose.Cells to create a named range that includes only visible rows after applying an AutoFilter. | Show how to iterate worksheet rows, detect hidden rows, build a comma‑separated filtered address, and assign it to Name.RefersTo. | Explain how to use the created named range in an Excel formula (e.g., SUM) and calculate the result with Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to build a workbook, apply an AutoFilter, collect the addresses of visible (non‑hidden) rows, combine them into a comma‑separated address, assign that address to a named range, use the name in a SUM formula, and save the file with Aspose.Cells for .NET.
    public class NamedRangeExcludingHiddenRows
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Apple");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("Cherry");
            sheet.Cells["B5"].PutValue(40);
            sheet.Cells["A6"].PutValue("Apple");
            sheet.Cells["B6"].PutValue(50);

            // Apply an AutoFilter to the header row covering columns A and B
            sheet.AutoFilter.Range = "A1:B6";

            // Filter to show only rows where Column A = "Apple"
            sheet.AutoFilter.AddFilter(0, "Apple");
            sheet.AutoFilter.Refresh();

            // Build an address that includes only the visible (non‑hidden) rows
            List<string> visibleRanges = new List<string>();
            int startRow = -1;
            int totalRows = sheet.Cells.MaxDataRow; // last row with data (zero‑based)

            for (int row = 1; row <= totalRows; row++) // start from row 2 (index 1) – data rows
            {
                if (!sheet.Cells.IsRowHidden(row))
                {
                    if (startRow == -1)
                        startRow = row; // begin a new visible block
                }
                else
                {
                    if (startRow != -1)
                    {
                        // End of a visible block – add its address
                        visibleRanges.Add($"'{sheet.Name}'!A{startRow + 1}:B{row}");
                        startRow = -1;
                    }
                }
            }

            // Add the last block if it ends at the bottom of the data
            if (startRow != -1)
                visibleRanges.Add($"'{sheet.Name}'!A{startRow + 1}:B{totalRows + 1}");

            // Join the individual blocks into a single address (comma‑separated)
            string filteredAddress = string.Join(",", visibleRanges);

            // Create a named range that refers to the filtered (visible) address
            int nameIndex = workbook.Worksheets.Names.Add("VisibleAppleRows");
            Name visibleName = workbook.Worksheets.Names[nameIndex];
            visibleName.RefersTo = "=" + filteredAddress; // e.g., =Sheet1!A2:B2,Sheet1!A4:B4,...

            // Demonstrate usage: sum the values in column B of the visible rows
            sheet.Cells["D1"].Formula = $"=SUM({visibleName.Text})";
            workbook.CalculateFormula();

            Console.WriteLine($"Named range '{visibleName.Text}' refers to: {visibleName.RefersTo}");
            Console.WriteLine($"Sum of visible rows (column B): {sheet.Cells["D1"].Value}");

            // Save the workbook
            workbook.Save("NamedRangeExcludingHiddenRows.xlsx");
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                NamedRangeExcludingHiddenRows.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
