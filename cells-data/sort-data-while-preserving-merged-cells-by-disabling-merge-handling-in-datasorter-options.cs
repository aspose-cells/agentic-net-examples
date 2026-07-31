// Title: Sort data with Aspose.Cells for .NET while preserving merged header cells
// Description: Demonstrates how to sort a worksheet range that contains a merged header row using Aspose.Cells for .NET. The example creates a workbook, adds a table (Category, Item, Quantity), merges the header (A1:C1), configures DataSorter with HasHeaders, sets multi‑column sort keys, defines a CellArea that includes the merged row, executes sorter.Sort, and saves the file. No extra option is needed because Aspose.Cells leaves merged cells untouched during sorting.
// Keywords: Aspose.Cells | .NET | C# | DataSorter | merged cells | preserve merged header | Excel sort merged cells | disable merge handling | sorting range | sample code | GitHub example
// Common Searches: Aspose.Cells sort range with merged header | C# keep merged cells when sorting Excel | DataSorter preserve merged cells Aspose | disable merge handling Aspose.Cells .NET | how to sort without breaking merged cells
// Developer Intent: The developer needs to reorder worksheet rows while ensuring that any merged header cells remain unchanged.
// Use Cases: Reorder a product catalog by category and quantity without unmerging the title row. | Sort financial statement rows while keeping the merged report title intact. | Organize inventory data with a multi‑column sort and a merged header for better presentation.
// AI Prompts: Generate C# code using Aspose.Cells to sort a range that includes a merged header row without unmerging it. | Explain how to configure DataSorter in Aspose.Cells for .NET to preserve merged cells during sorting. | Provide a complete example that sorts multiple columns while keeping merged cells intact, and save the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsSortingWithMergedCells
{
    // Demonstrates how to sort a worksheet range that contains a merged header row using Aspose.Cells for .NET. The example creates a workbook, adds a table (Category, Item, Quantity), merges the header (A1:C1), configures DataSorter with HasHeaders, sets multi‑column sort keys, defines a CellArea that includes the merged row, executes sorter.Sort, and saves the file. No extra option is needed because Aspose.Cells leaves merged cells untouched during sorting.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // Prepare sample data
            // -------------------------------------------------
            // Header row
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Item");
            cells["C1"].PutValue("Quantity");

            // Data rows
            cells["A2"].PutValue("Fruits");
            cells["B2"].PutValue("Apple");
            cells["C2"].PutValue(30);

            cells["A3"].PutValue("Fruits");
            cells["B3"].PutValue("Banana");
            cells["C3"].PutValue(20);

            cells["A4"].PutValue("Vegetables");
            cells["B4"].PutValue("Carrot");
            cells["C4"].PutValue(15);

            cells["A5"].PutValue("Vegetables");
            cells["B5"].PutValue("Tomato");
            cells["C5"].PutValue(25);

            // -------------------------------------------------
            // Merge the header cells (A1:C1) to demonstrate merged cells
            // -------------------------------------------------
            cells.Merge(0, 0, 1, 3); // Row 0, Column 0, 1 row, 3 columns

            // -------------------------------------------------
            // Configure DataSorter
            // -------------------------------------------------
            DataSorter sorter = workbook.DataSorter;

            // The range has a header (the merged header row)
            sorter.HasHeaders = true;

            // Sort by the "Category" column (index 0) then by "Quantity" column (index 2)
            sorter.AddKey(0, SortOrder.Ascending);
            sorter.AddKey(2, SortOrder.Descending);

            // -------------------------------------------------
            // Define the sort area (including the merged header row)
            // -------------------------------------------------
            CellArea sortArea = new CellArea
            {
                StartRow = 0,      // include merged header
                StartColumn = 0,
                EndRow = 5,        // rows 0‑5 (0‑4 data + header)
                EndColumn = 2
            };

            // -------------------------------------------------
            // Perform the sort
            // -------------------------------------------------
            // By default Aspose.Cells does not break merged cells when sorting.
            // No additional option is required to "disable merge handling".
            sorter.Sort(cells, sortArea);

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("SortedWithMergedHeader.xlsx");
        }
    }
}
