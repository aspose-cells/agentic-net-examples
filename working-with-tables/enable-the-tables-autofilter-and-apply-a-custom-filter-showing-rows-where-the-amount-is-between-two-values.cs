// Title: Aspose.Cells C# – Enable Table AutoFilter and Apply a Numeric Range Filter
// Description: Creates a workbook, adds a header and sample data, defines a ListObject (table) over A1:B6, turns on AutoFilter, and applies a custom filter on the "Amount" column to show rows with values between 100 and 500 before saving the file.
// Keywords: Aspose.Cells C# table auto filter | ListObject custom filter | filter rows between values | numeric range filter Aspose.Cells | Excel AutoFilter C# example | Aspose.Cells filter column by range | C# Excel table filter
// Common Searches: how to enable auto‑filter on a table using Aspose.Cells .NET | apply numeric range filter to ListObject column C# | Aspose.Cells filter rows where amount is between 100 and 500 | C# code for custom Excel table filter with Aspose.Cells | Aspose.Cells AutoFilter custom criteria example
// Developer Intent: Turn on AutoFilter for a ListObject and restrict displayed rows to those whose Amount lies within a specified numeric interval.
// Use Cases: Generate financial reports that only display transactions within a budget window. | Provide end‑users with a pre‑filtered Excel view that hides out‑of‑range values. | Automate data exports while excluding records that do not meet numeric criteria.
// AI Prompts: Write C# code with Aspose.Cells to add a table, enable AutoFilter, and filter column C for values between 200 and 800. | Explain how to modify the sample to use exclusive bounds and how to clear the custom filter afterward. | Create a snippet that applies separate custom filters on multiple columns of the same ListObject.

using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, adds a header and sample data, defines a ListObject (table) over A1:B6, turns on AutoFilter, and applies a custom filter on the "Amount" column to show rows with values between 100 and 500 before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add header row
        worksheet.Cells["A1"].PutValue("Item");
        worksheet.Cells["B1"].PutValue("Amount");

        // Populate sample data
        string[] items = { "A", "B", "C", "D", "E" };
        double[] amounts = { 50, 150, 300, 600, 400 };
        for (int i = 0; i < items.Length; i++)
        {
            worksheet.Cells[i + 1, 0].PutValue(items[i]);   // Column A
            worksheet.Cells[i + 1, 1].PutValue(amounts[i]); // Column B
        }

        // Create a ListObject (table) that covers the data range (A1:B6)
        int tableIndex = worksheet.ListObjects.Add(0, 0, items.Length, 1, true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Ensure the table includes all rows and columns
        table.Resize(0, 0, items.Length, 2, true);

        // Enable auto‑filter for the table
        table.HasAutoFilter = true;

        // Apply a custom filter on the "Amount" column (index 1) to show rows where
        // Amount is between 100 and 500 (inclusive)
        table.AutoFilter.Custom(
            fieldIndex: 1,
            operatorType1: FilterOperatorType.GreaterOrEqual,
            criteria1: 100,
            isAnd: true,
            operatorType2: FilterOperatorType.LessOrEqual,
            criteria2: 500);

        // Refresh the filter to hide non‑matching rows
        table.AutoFilter.Refresh();

        // Save the workbook
        workbook.Save("TableAutoFilterBetweenValues.xlsx", SaveFormat.Xlsx);
    }
}
