// Title: Clear all filters from an Aspose.Cells ListObject (table) in C# – restore full row visibility
// Description: This example creates a workbook, adds a ListObject with an auto‑filter, applies a filter on the "Category" column, then removes the filter using RemoveAutoFilter, checks HasAutoFilter and IsRowHidden to confirm that every row is visible, and saves the file.
// Keywords: Aspose.Cells Clear Table Filters C# | RemoveAutoFilter Aspose.Cells | Aspose.Cells ListObject reset filter | C# Aspose.Cells hide rows | Aspose.Cells table auto filter removal | .NET Aspose.Cells filter clear
// Common Searches: Aspose.Cells remove table filter C# | How to clear auto filter in Aspose.Cells ListObject | Reset filters in Aspose.Cells worksheet | Show hidden rows after filter Aspose.Cells | C# Aspose.Cells RemoveAutoFilter example
// Developer Intent: Remove the auto‑filter from a ListObject so that all rows become visible again.
// Use Cases: After programmatically filtering a table for analysis, clear the filter before saving or exporting the workbook. | Implement a "Reset Filters" button in a .NET application that uses Aspose.Cells to toggle table visibility for end users. | Reveal hidden rows prior to operations that require the complete dataset, such as chart generation or data export.
// AI Prompts: Generate C# code with Aspose.Cells that creates a ListObject, applies a filter on the "Category" column, then clears all filters using RemoveAutoFilter and verifies row visibility. | Explain the difference between ListObject.RemoveAutoFilter and manually clearing filter criteria in Aspose.Cells. | Provide a step‑by‑step guide to add a "Reset Filters" feature in a WinForms app using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// This example creates a workbook, adds a ListObject with an auto‑filter, applies a filter on the "Category" column, then removes the filter using RemoveAutoFilter, checks HasAutoFilter and IsRowHidden to confirm that every row is visible, and saves the file.
class ClearTableFilters
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with a header row
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Category");
        sheet.Cells["C1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Food");
        sheet.Cells["C2"].PutValue(100);
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Drink");
        sheet.Cells["C3"].PutValue(50);
        sheet.Cells["A4"].PutValue(3);
        sheet.Cells["B4"].PutValue("Food");
        sheet.Cells["C4"].PutValue(150);

        // Add a ListObject (table) that includes an auto‑filter by default
        int tableIndex = sheet.ListObjects.Add(0, 0, 3, 2, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Apply a filter to the "Category" column (index 1) to show only "Food"
        table.AutoFilter.AddFilter(1, "Food");
        table.AutoFilter.Refresh();

        // Demonstrate that rows not matching the filter are hidden
        Console.WriteLine("Row 3 hidden after filter: " + sheet.Cells.IsRowHidden(2)); // Row index 2 = Excel row 3

        // Clear all filters from the table, restoring full visibility
        table.RemoveAutoFilter();

        // Verify that the auto‑filter has been removed and all rows are visible
        Console.WriteLine("HasAutoFilter after removal: " + table.HasAutoFilter);
        Console.WriteLine("Row 3 hidden after clearing filters: " + sheet.Cells.IsRowHidden(2));

        // Save the workbook
        workbook.Save("ClearTableFilters.xlsx");
    }
}
