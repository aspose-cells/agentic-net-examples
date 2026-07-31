// Title: C# – Create a Named Range Excluding AutoFilter‑Hidden Rows for SUM and AVERAGE with Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a workbook, fills columns A and B with sample data, applies an AutoFilter to show rows where Quantity > 8, gathers the addresses of the visible cells in column B, defines a named range "VisibleQuantities" that references only those cells, and uses the range in SUM and AVERAGE formulas before saving the file.
// Keywords: Aspose.Cells named range visible rows | C# filter hidden rows Aspose | exclude autofilter rows named range | SUM AVERAGE visible cells Aspose.Cells | .NET Excel named range filter | dynamic named range Aspose.Cells
// Common Searches: Aspose.Cells create named range from visible cells | C# sum of filtered rows using Aspose.Cells | how to ignore hidden rows in Excel formulas with Aspose | define named range for AutoFilter results .NET | calculate average of visible rows Aspose.Cells
// Developer Intent: Generate a named range that contains only the cells visible after an AutoFilter and use it for statistical calculations in C#.
// Use Cases: Compute total and average of quantities for items that meet a filter condition. | Supply a filtered‑only data series to charts or pivot tables. | Programmatically retrieve the address of visible cells for reporting or further processing.
// AI Prompts: Write C# code using Aspose.Cells to build a named range that excludes rows hidden by an AutoFilter and then calculate SUM and AVERAGE. | Explain how to collect the addresses of visible cells after applying a filter and assign them to a named range in Aspose.Cells. | Show how to reference a dynamically created named range in a formula and obtain its Range object for additional manipulation.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This Aspose.Cells for .NET example creates a workbook, fills columns A and B with sample data, applies an AutoFilter to show rows where Quantity > 8, gathers the addresses of the visible cells in column B, defines a named range "VisibleQuantities" that references only those cells, and uses the range in SUM and AVERAGE formulas before saving the file.
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
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Quantity");
                for (int i = 2; i <= 15; i++)
                {
                    cells[$"A{i}"].PutValue($"Item {i - 1}");
                    cells[$"B{i}"].PutValue(i - 1); // 1..14
                }

                // Apply an AutoFilter to the header row covering columns A and B
                sheet.AutoFilter.Range = "A1:B15";

                // Filter: show only rows where Quantity > 8
                sheet.AutoFilter.Custom(1, FilterOperatorType.GreaterThan, 8);
                sheet.AutoFilter.Refresh();

                // Build address string of visible cells in column B
                string visibleAddress = "";
                int maxRow = cells.MaxDataRow;
                for (int row = 1; row <= maxRow; row++) // zero‑based index, row 1 = Excel row 2
                {
                    if (!sheet.Cells.IsRowHidden(row))
                    {
                        string cellAddr = cells[row, 1].Name; // column B
                        visibleAddress += cellAddr + ",";
                    }
                }
                if (visibleAddress.EndsWith(","))
                    visibleAddress = visibleAddress.Substring(0, visibleAddress.Length - 1);

                // Create a named range that refers only to the visible cells
                int nameIdx = workbook.Worksheets.Names.Add("VisibleQuantities");
                Aspose.Cells.Name visibleName = workbook.Worksheets.Names[nameIdx];
                visibleName.RefersTo = $"={sheet.Name}!{visibleAddress}";

                // Use the named range in formulas
                cells["D1"].PutValue("Sum of Visible Quantities:");
                cells["E1"].Formula = "=SUM(VisibleQuantities)";

                cells["D2"].PutValue("Average of Visible Quantities:");
                cells["E2"].Formula = "=AVERAGE(VisibleQuantities)";

                // Retrieve the range object via GetRange
                Aspose.Cells.Range rng = visibleName.GetRange();
                Console.WriteLine($"Named range '{visibleName.Text}' refers to: {rng.Address}");

                // Save the workbook
                string outputPath = "NamedRangeExcludingFilteredRows.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the example
    public class Program
    {
        public static void Main()
        {
            NamedRangeExcludingFilteredRows.Run();
        }
    }
}
