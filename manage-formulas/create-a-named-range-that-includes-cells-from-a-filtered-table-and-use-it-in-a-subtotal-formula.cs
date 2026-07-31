// Title: Define a Named Range for a Filtered Table and Apply Subtotal in Aspose.Cells for .NET
// Description: Shows how to create a workbook, fill it with sales data, apply an AutoFilter, define a named range that includes the filtered rows, convert the range to a CellArea, and use the Subtotal method to group by Region, sum the Sales column, add page breaks, place the summary below the data, and save the file.
// Keywords: Aspose.Cells | C# | .NET | named range | AutoFilter | Subtotal method | group by column | sum function | page break | filtered table | Excel automation
// Common Searches: Aspose.Cells create named range for filtered data | How to use Subtotal with AutoFilter in Aspose.Cells | C# subtotal grouped by column after applying filter | Define named range that includes hidden rows Aspose.Cells | Add page breaks with Subtotal method in .NET
// Developer Intent: The developer wants to define a named range that spans a filtered table and then generate subtotal totals based on that range.
// Use Cases: Produce region‑wise sales totals after applying a filter | Create printable reports with automatic page breaks between groups | Reuse the same named range in charts, pivot tables, or other formulas
// AI Prompts: Generate code to subtotal multiple columns (e.g., Quantity and Sales) using the same named range. | Add a grand‑total row that aggregates the subtotals for the filtered dataset. | Explain how to make the named range automatically adjust when the filter criteria change.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, fill it with sales data, apply an AutoFilter, define a named range that includes the filtered rows, convert the range to a CellArea, and use the Subtotal method to group by Region, sum the Sales column, add page breaks, place the summary below the data, and save the file.
    public class NamedRangeWithFilteredTableSubtotal
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (header + 6 rows)
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            object[,] data = new object[,]
            {
                {"North", "Widget", 5000},
                {"North", "Gadget", 3000},
                {"South", "Widget", 6000},
                {"South", "Gadget", 4000},
                {"West",  "Widget", 4500},
                {"West",  "Gadget", 3500}
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    cells[r + 1, c].PutValue(data[r, c]);
                }
            }

            // Define the area that will be filtered (A1:C7)
            CellArea filterArea = CellArea.CreateCellArea("A1", "C7");

            // Build the reference string for the area (e.g., "A1:C7")
            string startRef = CellsHelper.CellIndexToName(filterArea.StartRow, filterArea.StartColumn);
            string endRef = CellsHelper.CellIndexToName(filterArea.EndRow, filterArea.EndColumn);
            string filterRef = $"{startRef}:{endRef}";

            // Apply AutoFilter to the defined area
            sheet.AutoFilter.Range = filterRef;

            // Apply a filter: show only rows where Region = "North"
            sheet.AutoFilter.Filter(0, "North"); // column index 0 corresponds to "Region"

            // Create a named range that refers to the same area (including hidden rows)
            int nameIdx = workbook.Worksheets.Names.Add("FilteredData");
            workbook.Worksheets.Names[nameIdx].RefersTo = $"={sheet.Name}!{filterRef}";

            // Retrieve the Range object from the named range
            Name namedRange = workbook.Worksheets.Names["FilteredData"];
            Aspose.Cells.Range range = namedRange.GetRange();

            // Convert the Range to a CellArea for the Subtotal method
            CellArea subtotalArea = CellArea.CreateCellArea(
                range.FirstRow,
                range.FirstColumn,
                range.FirstRow + range.RowCount - 1,
                range.FirstColumn + range.ColumnCount - 1);

            // Add subtotals:
            // - Group by the first column (Region)
            // - Use SUM function on the Sales column (index 2)
            // - Replace existing subtotals, add page breaks, place summary below data
            sheet.Cells.Subtotal(
                subtotalArea,
                0,                                   // group by column 0 (Region)
                ConsolidationFunction.Sum,           // SUM function
                new int[] { 2 },                     // apply subtotal to column 2 (Sales)
                true,                                // replace existing subtotals
                true,                                // add page breaks between groups
                true);                               // place summary below data

            // Save the workbook
            try
            {
                workbook.Save("NamedRangeFilteredSubtotal.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
