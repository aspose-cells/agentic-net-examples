// Title: Apply a custom descending text sort to a worksheet column with Aspose.Cells for .NET
// AI Prompts: Configure Aspose.Cells DataSorter to sort column A in descending alphabetical order while treating values as text and ignoring case. | Create a C# routine that sorts a range containing a header row using Aspose.Cells, with SortOrder.Descending and SortAsNumber set to false. | Generate an Excel file where the 'Category' column is ordered Z‑to‑A using Aspose.Cells DataSorter and save it as .xlsx.
// Common Searches: Aspose.Cells sort column as text descending with header row C# | DataSorter descending alphabetical sort example Aspose.Cells .NET | How to perform case‑insensitive text sort on Excel column using Aspose.Cells | C# Aspose.Cells custom sort range descending string values | Sorting Excel data by string values in reverse order with Aspose.Cells
// Tags: DataSorter descending text sort Aspose.Cells | custom column ordering with Aspose.Cells | sort Excel range with header Aspose.Cells | case‑insensitive alphabetical sorting .NET | save sorted worksheet as .xlsx Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsCustomTextSort
{
    // Demonstrates using Aspose.Cells' DataSorter to sort column A of a worksheet in descending alphabetical order, treating values as text, handling a header row, and saving the result to an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including a header)
            cells["A1"].PutValue("Category");
            cells["A2"].PutValue("Banana");
            cells["A3"].PutValue("Apple");
            cells["A4"].PutValue("Cherry");
            cells["A5"].PutValue("Date");

            // Configure the DataSorter
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;               // First row is a header
            sorter.CaseSensitive = false;           // Case‑insensitive sorting
            sorter.SortAsNumber = false;            // Treat values as text, not numbers

            // Set the first sort key to column A (index 0) with descending order
            sorter.Key1 = 0;                         // Column A
            sorter.Order1 = SortOrder.Descending;   // Descending (Z → A)

            // Define the range to sort (including header row)
            CellArea sortArea = CellArea.CreateCellArea("A1", "A5");

            // Perform the sort
            sorter.Sort(cells, sortArea);

            // Save the result
            workbook.Save("CustomTextDescendingSort.xlsx");
        }
    }
}
