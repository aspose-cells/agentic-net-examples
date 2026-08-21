// Title: Aspose.Cells .NET: Create a Named Range from a Filtered Table and Apply Subtotal
// Description: Demonstrates how to build a workbook, add sample sales data, set an AutoFilter on A1:C6, capture the filtered CellArea, define a named range "FilteredData", and use Cells.Subtotal to group by Category and sum Sales with page breaks and a summary row.
// Keywords: Aspose.Cells named range from AutoFilter | C# subtotal method Aspose.Cells | filter area named range .NET | grouped subtotals Aspose.Cells | AutoFilter Subtotal example | Aspose.Cells Subtotal function | create named range programmatically | Excel subtotal with Aspose.Cells
// Common Searches: Aspose.Cells create named range after applying AutoFilter | How to use Subtotal method on filtered data in C# | Define named range for filtered rows Aspose.Cells .NET | Apply SUM subtotal by category using Aspose.Cells | Add page breaks with Subtotal in Aspose.Cells
// Developer Intent: Generate a named range that references the filtered portion of a table and use it to produce grouped SUM subtotals via the Cells.Subtotal API.
// Use Cases: Produce category‑wise sales totals after the user filters data. | Create printable reports that insert page breaks between each group. | Reuse the "FilteredData" range in charts, formulas, or pivot tables.
// AI Prompts: Write C# code with Aspose.Cells to define a named range from an AutoFilter area and apply a SUM subtotal on the Sales column grouped by Category. | Extend the example to add a COUNT subtotal for the Product column while preserving the existing named range. | Show how to reference the "FilteredData" named range in a formula on another worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to build a workbook, add sample sales data, set an AutoFilter on A1:C6, capture the filtered CellArea, define a named range "FilteredData", and use Cells.Subtotal to group by Category and sum Sales with page breaks and a summary row.
    public class NamedRangeFilteredSubtotalDemo
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

            // Populate sample data (including header)
            // Header: Category, Product, Sales
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            object[,] data = new object[,]
            {
                {"North", "Widget", 5000},
                {"North", "Gadget", 3000},
                {"South", "Widget", 6000},
                {"South", "Gadget", 4000},
                {"West",  "Widget", 4500}
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                cells[i + 1, 0].PutValue(data[i, 0]); // Column A
                cells[i + 1, 1].PutValue(data[i, 1]); // Column B
                cells[i + 1, 2].PutValue(data[i, 2]); // Column C
            }

            // Apply an AutoFilter to the data range (A1:C6)
            // SetRange(startRow, startColumn, totalColumns)
            // startRow = 0 (row 1), startColumn = 0 (column A), totalColumns = 3 (A,B,C)
            sheet.AutoFilter.SetRange(0, 0, 3);

            // Retrieve the actual CellArea where the filter is applied
            CellArea filterArea = sheet.AutoFilter.GetCellArea();

            // Create a named range that refers to the filtered area
            int rowCount = filterArea.EndRow - filterArea.StartRow + 1;
            int colCount = filterArea.EndColumn - filterArea.StartColumn + 1;
            Aspose.Cells.Range filteredRange = cells.CreateRange(filterArea.StartRow, filterArea.StartColumn, rowCount, colCount);
            filteredRange.Name = "FilteredData";

            // Use the Subtotal method on the same area
            // Group by the first column (Category) -> index 0
            // Apply SUM to the Sales column (index 2)
            // Replace existing subtotals, add page breaks, place summary below data
            cells.Subtotal(
                filterArea,
                0,                                 // groupBy column index
                ConsolidationFunction.Sum,         // subtotal function
                new int[] { 2 },                   // columns to subtotal
                true,                              // replace existing subtotals
                true,                              // add page breaks between groups
                true                               // place summary below data
            );

            // Save the workbook
            string outputPath = "NamedRangeFilteredSubtotalDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
