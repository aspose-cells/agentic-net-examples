// Title: Filter a ListObject to display only rows where the Region column equals 'East' using Aspose.Cells in C#
// AI Prompts: Create a worksheet, add a ListObject covering a data range, enable AutoFilter, and filter the Region column for the value 'East' with Aspose.Cells in C#. | Use the Aspose.Cells ListObject.AutoFilter API to apply a criteria filter on column index 0, refresh the filter, and save the workbook.
// Common Searches: Aspose.Cells .NET how to filter a ListObject by column value | C# example of using AutoFilter on an Excel table with Aspose.Cells | Show only rows where Region = East in an Aspose.Cells workbook | Apply a criteria filter to a ListObject column using Aspose.Cells API
// Tags: Aspose.Cells ListObject AutoFilter | filter rows by column value Aspose.Cells | C# apply ListObject criteria filter | Excel table region filter Aspose.Cells | Aspose.Cells workbook table filtering

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // The example creates a workbook, populates Region and Sales data, defines a ListObject (table) over the range, enables AutoFilter, applies a filter on the Region column to keep only rows with the value "East", refreshes the filter, and saves the file as ListObjectRegionFilterDemo.xlsx.
    public class ListObjectRegionFilterDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
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

            // Populate sample data with a header row
            // Column A: Region, Column B: Sales
            worksheet.Cells["A1"].PutValue("Region");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["A2"].PutValue("East");
            worksheet.Cells["B2"].PutValue(1200);
            worksheet.Cells["A3"].PutValue("West");
            worksheet.Cells["B3"].PutValue(950);
            worksheet.Cells["A4"].PutValue("East");
            worksheet.Cells["B4"].PutValue(800);
            worksheet.Cells["A5"].PutValue("North");
            worksheet.Cells["B5"].PutValue(670);

            // Add a ListObject (table) that covers the data range (including header)
            // Parameters: startRow, startColumn, endRow, endColumn, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject listObject = worksheet.ListObjects[tableIndex];

            // Ensure the table has an AutoFilter enabled
            listObject.HasAutoFilter = true;

            // Apply filter on the "Region" column (index 0) to show only rows where Region = "East"
            listObject.AutoFilter.Filter(0, "East");
            // Refresh the filter to apply the changes
            listObject.AutoFilter.Refresh();

            // Save the workbook
            workbook.Save("ListObjectRegionFilterDemo.xlsx");
        }
    }
}
