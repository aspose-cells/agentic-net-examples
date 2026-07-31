// Title: Remove All Table Filters with Aspose.Cells ListObject in C#
// Description: Shows how to create a workbook, add a ListObject (Excel table) with an auto‑filter, apply a column filter, then clear every filter using ListObject.RemoveAutoFilter and save the result.
// Keywords: Aspose.Cells | C# | ListObject | RemoveAutoFilter | clear table filters | reset Excel table filters | programmatic filter removal | .NET Excel automation | Excel auto‑filter | table filter reset
// Common Searches: Aspose.Cells remove all filters from a table | Clear ListObject auto‑filter C# | Reset Excel table filters programmatically .NET | How to delete table filters with Aspose.Cells | RemoveAutoFilter example for Aspose.Cells
// Developer Intent: Clear any active auto‑filters on a ListObject so the full dataset is visible.
// Use Cases: After performing a filtered analysis, clear filters before saving to preserve the complete data set. | In automated reporting, ensure exported Excel files contain every row by resetting table filters programmatically. | When reusing a template that may have pre‑defined filters, reset tables to their default state before populating new data.
// AI Prompts: Generate C# code that detects active filters on a ListObject and removes them using Aspose.Cells. | Provide a snippet to clear filters for all ListObjects on a worksheet with Aspose.Cells for .NET. | Explain the difference between ListObject.RemoveAutoFilter and clearing individual filter criteria in Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Tables;

// Shows how to create a workbook, add a ListObject (Excel table) with an auto‑filter, apply a column filter, then clear every filter using ListObject.RemoveAutoFilter and save the result.
class ClearTableFilters
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with a header row
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Category");
        worksheet.Cells["C1"].PutValue("Amount");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("Food");
        worksheet.Cells["C2"].PutValue(100);
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue("Drink");
        worksheet.Cells["C3"].PutValue(50);
        worksheet.Cells["A4"].PutValue(3);
        worksheet.Cells["B4"].PutValue("Food");
        worksheet.Cells["C4"].PutValue(200);

        // Add a ListObject (table) that includes an auto‑filter by default
        int listIndex = worksheet.ListObjects.Add(0, 0, 3, 2, true);
        ListObject listObj = worksheet.ListObjects[listIndex];

        // Apply a filter on the "Category" column (index 1) to show only "Food"
        listObj.AutoFilter.AddFilter(1, "Food");
        listObj.AutoFilter.Refresh();

        // Clear all filters from the table, restoring full visibility
        listObj.RemoveAutoFilter();

        // Save the workbook
        workbook.Save("ClearAllFiltersDemo.xlsx");
    }
}
