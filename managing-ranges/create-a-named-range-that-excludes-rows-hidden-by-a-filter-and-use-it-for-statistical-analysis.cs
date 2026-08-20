// Title: Create a Named Range Excluding Filtered Rows for Statistical Analysis with Aspose.Cells .NET
// Description: Demonstrates how to build a workbook, apply an AutoFilter, collect only the visible cells in column B, define a named range that references those cells, and use the AVERAGE function to compute a statistic on the filtered data. The example calculates the formula, outputs the result, and saves the file.
// Keywords: Aspose.Cells named range filtered rows | C# visible rows union range | average visible cells Aspose.Cells | exclude hidden rows Aspose.Cells | AutoFilter visible range .NET | statistical analysis Aspose.Cells
// Common Searches: Aspose.Cells create named range from visible rows | C# average of filtered data using named range | how to exclude hidden rows in Aspose.Cells range | build union range of visible cells Aspose.Cells | calculate AVERAGE on filtered rows Aspose.Cells
// Developer Intent: Generate a named range that references only non‑hidden rows after filtering and apply it in aggregate formulas.
// Use Cases: Define a named range for the visible Amount values after filtering by Category and compute the average with AVERAGE. | Reuse the visible‑only named range for other aggregates such as SUM, COUNT, or MAX on filtered data. | Save the workbook so external Excel users can work with the pre‑filtered named range without additional code.
// AI Prompts: Write C# code with Aspose.Cells that creates a union range of visible cells in column B after applying an AutoFilter and assigns it to a named range. | Show how to apply an AutoFilter, build a named range that excludes hidden rows, and calculate the AVERAGE of that range using Aspose.Cells. | Explain how to retrieve the result of a formula that uses a named range representing only visible rows in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to build a workbook, apply an AutoFilter, collect only the visible cells in column B, define a named range that references those cells, and use the AVERAGE function to compute a statistic on the filtered data. The example calculates the formula, outputs the result, and saves the file.
    public class NamedRangeExcludingFilteredRows
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (header + values)
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Amount");
                cells["A2"].PutValue("Food");      cells["B2"].PutValue(120);
                cells["A3"].PutValue("Transport"); cells["B3"].PutValue(80);
                cells["A4"].PutValue("Food");      cells["B4"].PutValue(150);
                cells["A5"].PutValue("Travel");    cells["B5"].PutValue(200);
                cells["A6"].PutValue("Food");      cells["B6"].PutValue(90);

                // Apply an AutoFilter to the header row (A1:B6)
                sheet.AutoFilter.Range = "A1:B6";

                // Filter to show only rows where Category = "Food"
                sheet.AutoFilter.AddFilter(0, "Food");
                sheet.AutoFilter.Refresh(); // Hide non‑matching rows

                // Build a union range string that includes only visible rows in column B (Amount)
                // Example result: "B2,B4,B6"
                System.Text.StringBuilder visibleRangeBuilder = new System.Text.StringBuilder();
                int maxRow = cells.MaxDataRow; // last row with data (zero‑based)
                for (int row = 1; row <= maxRow; row++) // start from 1 because row 0 is header
                {
                    if (!sheet.Cells.IsRowHidden(row))
                    {
                        if (visibleRangeBuilder.Length > 0)
                            visibleRangeBuilder.Append(",");

                        // Excel rows are 1‑based, so add 1 to the zero‑based index
                        visibleRangeBuilder.Append($"B{row + 1}");
                    }
                }

                // Create a named range that refers only to the visible cells
                int nameIndex = workbook.Worksheets.Names.Add("VisibleAmounts");
                Name visibleName = workbook.Worksheets.Names[nameIndex];
                visibleName.RefersTo = $"={sheet.Name}!{visibleRangeBuilder}";

                // Use the named range in a statistical formula (AVERAGE)
                cells["D1"].PutValue("Average of Visible Amounts");
                cells["D2"].Formula = "=AVERAGE(VisibleAmounts)";

                // Calculate formulas so the result is stored
                workbook.CalculateFormula();

                // Display the calculated average in console
                Console.WriteLine("Average of visible amounts: " + cells["D2"].StringValue);

                // Save the workbook
                workbook.Save("NamedRangeExcludingFilteredRows.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            NamedRangeExcludingFilteredRows.Run();
        }
    }
}
