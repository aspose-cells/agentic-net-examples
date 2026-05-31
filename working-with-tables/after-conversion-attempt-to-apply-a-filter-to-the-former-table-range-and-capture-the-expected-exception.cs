using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data (including a header row)
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("Alice");
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue("Bob");
        worksheet.Cells["A4"].PutValue(3);
        worksheet.Cells["B4"].PutValue("Charlie");

        // Add a ListObject (table) that covers the data range
        int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
        ListObject listObject = worksheet.ListObjects[tableIndex];

        // Enable auto‑filter for the table and apply a simple filter
        listObject.HasAutoFilter = true;
        listObject.AutoFilter.Custom(0, FilterOperatorType.GreaterOrEqual, 2);
        listObject.AutoFilter.Refresh();

        // Convert the table back to a normal range
        listObject.ConvertToRange();

        // After conversion the ListObject is removed from the worksheet.
        // Attempt to apply a filter using the former ListObject reference.
        // This should raise an exception because the object is no longer part of the sheet.
        try
        {
            // The Filter method is obsolete but still callable; it will fail here.
            listObject.Filter();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Expected exception caught: " + ex.Message);
        }

        // Demonstrate that the worksheet can still apply a filter to the same cell area.
        try
        {
            CellArea filterArea = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 0,   // Column A
                EndRow = 3,        // Row 4
                EndColumn = 1      // Column B
            };
            worksheet.Filter(filterArea);
            Console.WriteLine("Worksheet.Filter applied successfully after conversion.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Worksheet.Filter exception: " + ex.Message);
        }

        // Save the workbook to verify the final state
        workbook.Save("TableConversionFilterDemo.xlsx");
    }
}