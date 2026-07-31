// Title: C# – Create a Named Range of Visible Rows (Exclude Hidden) Using AutoFilter Address in Aspose.Cells
// Description: Creates a workbook, applies an AutoFilter, builds a comma‑separated address of only the visible rows, adds a named range that points to that address, uses the name in a SUM formula, and saves the file.
// Keywords: Aspose.Cells | C# named range visible rows | exclude hidden rows Aspose.Cells | AutoFilter address | filtered range named range | Sum visible rows Aspose.Cells | Aspose.Cells .NET example
// Common Searches: Aspose.Cells create named range for filtered rows | C# named range only visible rows after AutoFilter | How to exclude hidden rows from named range in Aspose.Cells | Sum visible cells using named range Aspose.Cells | Build address string for visible rows Aspose.Cells
// Developer Intent: Generate a named range that references only the rows visible after an AutoFilter, omitting hidden rows.
// Use Cases: Calculate the total of a filtered column without counting hidden rows. | Provide a chart data source that displays only rows meeting the filter criteria. | Apply conditional formatting or data validation exclusively to visible rows. | Export a subset of data defined by the visible‑row named range. | Reference the visible rows from external tools or scripts via the named range.
// AI Prompts: Write C# Aspose.Cells code that creates a named range containing only non‑hidden rows after applying an AutoFilter. | Show how to generate a comma‑separated address of each visible row and assign it to Name.RefersTo. | Explain how to use the created named range in a worksheet formula such as SUM to total visible values. | Provide error handling for cases where no rows are visible after filtering. | Demonstrate how to reuse the named range as a chart data source in Aspose.Cells.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, applies an AutoFilter, builds a comma‑separated address of only the visible rows, adds a named range that points to that address, uses the name in a SUM formula, and saves the file.
    public class NamedRangeExcludingHiddenRows
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (header + rows)
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

                // Apply an AutoFilter to the whole data range (including header)
                sheet.AutoFilter.Range = "A1:B6";

                // Filter to show only rows where Category = "Apple"
                sheet.AutoFilter.AddFilter(0, "Apple");
                sheet.AutoFilter.Refresh();

                // Get the area covered by the AutoFilter (including header)
                CellArea filterArea = sheet.AutoFilter.GetCellArea(true);

                // Build an address that includes only the visible (non‑hidden) rows
                StringBuilder visibleAddress = new StringBuilder();
                for (int row = filterArea.StartRow + 1; row <= filterArea.EndRow; row++) // skip header row
                {
                    if (!sheet.Cells.IsRowHidden(row))
                    {
                        // Convert column indexes to letters
                        string startColLetter = CellsHelper.ColumnIndexToName(filterArea.StartColumn);
                        string endColLetter = CellsHelper.ColumnIndexToName(filterArea.EndColumn);
                        // Build address for the current row (e.g., A2:B2)
                        string rowAddress = $"{startColLetter}{row + 1}:{endColLetter}{row + 1}";
                        if (visibleAddress.Length > 0)
                            visibleAddress.Append(",");
                        visibleAddress.Append(rowAddress);
                    }
                }

                // Create a named range that refers to the visible rows only
                int nameIndex = workbook.Worksheets.Names.Add("VisibleAppleRows");
                Name visibleRangeName = workbook.Worksheets.Names[nameIndex];
                // RefersTo must start with '=' and include the sheet name
                visibleRangeName.RefersTo = $"={sheet.Name}!{visibleAddress}";

                // Demonstrate that the named range works in a formula
                sheet.Cells["C1"].Formula = $"=SUM({visibleRangeName.Text})";
                workbook.CalculateFormula();

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
