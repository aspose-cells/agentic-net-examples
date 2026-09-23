// Title: Enable auto‑filter on an Aspose.Cells ListObject and filter rows that contain a specific substring using C#
// AI Prompts: Write C# code that creates a workbook, adds a ListObject table, enables its header row, and applies an AutoFilter with a '*apple*' wildcard to show only rows where the first column contains 'apple'. | Show how to set a custom text criteria with wildcards on a specific column of an Aspose.Cells table and then save the filtered result as an .xlsx file. | Demonstrate applying a table style, turning on auto‑filter, and using the AutoFilter.Filter method to perform substring matching in an Excel worksheet with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# how to filter a ListObject by text containing a substring | apply wildcard auto filter to Excel table using Aspose.Cells .NET | filter rows in an Aspose.Cells table where column value includes 'apple' | enable auto filter on Aspose.Cells table and save filtered workbook
// Tags: Aspose.Cells auto filter ListObject | C# filter Excel table by substring | Aspose.Cells wildcard text criteria | apply table style Aspose.Cells | save filtered workbook .xlsx Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, defines a ListObject covering A1:B5, enables the header row, applies a table style, and uses AutoFilter.Filter with the wildcard '*apple*' to display only rows whose Name contains 'apple'. The filtered workbook is saved as FilteredTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue("Fruit");
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue("Fruit");
            worksheet.Cells["A4"].PutValue("Carrot");
            worksheet.Cells["B4"].PutValue("Vegetable");
            worksheet.Cells["A5"].PutValue("Pineapple");
            worksheet.Cells["B5"].PutValue("Fruit");

            // Define the range for the table (A1:B5)
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 5;     // rows 1‑5
            int totalColumns = 2;  // columns A‑B

            // Add a ListObject (table) with headers
            int tableIndex = worksheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.ShowHeaderRow = true;
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Apply a filter to show rows where "Name" contains "apple"
            // Using wildcard pattern "*apple*"
            table.AutoFilter.Filter(0, "*apple*");

            // Save the workbook
            string outputPath = "FilteredTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
