// Title: Custom Sort Worksheet by Predefined String List (High‑Medium‑Low) with Aspose.Cells for .NET
// Description: Demonstrates how to use Aspose.Cells' DataSorter to order rows based on a user‑defined list (High, Medium, Low). The example creates a workbook, adds a two‑column table, sets the custom list, applies the sort with headers, prints the result, and saves the file.
// Keywords: Aspose.Cells custom list sorting | DataSorter AddKey string list | C# sort Excel range by priority | custom order sort Aspose.Cells .NET | Excel custom priority sort code
// Common Searches: Aspose.Cells sort by custom list .NET | C# sort worksheet using predefined string order | How to apply custom priority sorting in Aspose.Cells | DataSorter custom order example | Sort Excel column with High Medium Low list
// Developer Intent: Arrange worksheet rows according to a specific priority sequence defined by a string list.
// Use Cases: Generate a task report where items appear as High → Medium → Low. | Prepare product data in a business‑defined order before feeding a pivot table. | Export CSV files that must follow a custom sequence mandated by a client.
// AI Prompts: Write C# code that sorts an Excel range with Aspose.Cells using a custom string list and respects header rows. | Explain how to change the custom list at runtime and re‑apply the sort on an existing worksheet. | Provide a guide to sort multiple columns, each with its own custom order, using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomSortExample
{
    // Demonstrates how to use Aspose.Cells' DataSorter to order rows based on a user‑defined list (High, Medium, Low). The example creates a workbook, adds a two‑column table, sets the custom list, applies the sort with headers, prints the result, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data with a header row
            // Column A: Item, Column B: Priority (to be sorted by custom order)
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Priority");

            cells["A2"].PutValue("Task 1");
            cells["B2"].PutValue("Medium");

            cells["A3"].PutValue("Task 2");
            cells["B3"].PutValue("Low");

            cells["A4"].PutValue("Task 3");
            cells["B4"].PutValue("High");

            cells["A5"].PutValue("Task 4");
            cells["B5"].PutValue("Medium");

            // Define the custom sort list for the Priority column
            // The order will be: High -> Medium -> Low
            string customList = "High,Medium,Low";

            // Get the DataSorter from the workbook
            DataSorter sorter = workbook.DataSorter;

            // Indicate that the range contains headers
            sorter.HasHeaders = true;

            // Add a sort key for column B (index 1) using the custom list
            sorter.AddKey(1, SortOrder.Ascending, customList);

            // Define the range to sort (including header row)
            CellArea sortArea = CellArea.CreateCellArea("A1", "B5");

            // Perform the sort
            sorter.Sort(worksheet.Cells, sortArea);

            // Output the sorted result to the console
            Console.WriteLine("Sorted data based on custom priority order:");
            for (int row = 1; row <= 5; row++) // start from row 2 (index 1) to skip header
            {
                string item = cells[row, 0].StringValue;
                string priority = cells[row, 1].StringValue;
                Console.WriteLine($"{item}: {priority}");
            }

            // Save the workbook to a file
            workbook.Save("CustomSortedData.xlsx");
        }
    }
}
