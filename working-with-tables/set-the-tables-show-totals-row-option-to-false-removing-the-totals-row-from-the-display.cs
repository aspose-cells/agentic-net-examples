// Title: Hide the Totals Row of an Aspose.Cells ListObject (Table) in C# (.NET)
// Description: C# example that creates a workbook, adds sample data, defines a ListObject covering A1:B4, and hides the totals row by setting ListObject.ShowTotals to false before saving as RemoveTotalsRowDemo.xlsx.
// Keywords: Aspose.Cells | C# | ListObject ShowTotals | hide totals row | remove table totals row | disable totals row Excel | Aspose.Cells .NET table options | Excel ListObject hide totals | Aspose.Cells Table ShowTotals false
// Common Searches: Aspose.Cells hide totals row C# | Set ListObject ShowTotals to false | Remove totals row from Excel table using Aspose.Cells | How to disable totals row in Aspose.Cells .NET | Aspose.Cells ListObject ShowTotals property example
// Developer Intent: Programmatically hide the totals row of an Excel table (ListObject) by setting the ShowTotals property to false using Aspose.Cells for .NET.
// Use Cases: Generate Excel reports without a totals row for a cleaner layout. | Toggle the visibility of the totals row based on user preferences at runtime. | Create a workbook, add data, define a table, and ensure the totals row is not displayed before saving.
// AI Prompts: Write C# code with Aspose.Cells that creates a table and hides its totals row. | Explain how the ListObject.ShowTotals property works and how to set it to false. | Provide a snippet that adds a ListObject, enables the totals row, then disables it by setting ShowTotals to false.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// C# example that creates a workbook, adds sample data, defines a ListObject covering A1:B4, and hides the totals row by setting ListObject.ShowTotals to false before saving as RemoveTotalsRowDemo.xlsx.
class RemoveTotalsRowDemo
{
    static void Main(string[] args)
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
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data for the table
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Price");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["B3"].PutValue(15);
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B4"].PutValue(8);

        // Create a ListObject (table) covering the data range
        int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // (Optional) Enable the totals row first
        table.ShowTotals = true;

        // Hide the totals row by setting ShowTotals to false
        table.ShowTotals = false;

        // Save the workbook
        workbook.Save("RemoveTotalsRowDemo.xlsx");
        Console.WriteLine("Workbook saved as RemoveTotalsRowDemo.xlsx");
    }
}
